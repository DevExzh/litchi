//! OPC singleton graph lifecycle for an XLSX workbook Data Model.

use std::collections::HashMap;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::{Reader, events::Event};

use crate::error::{Error, Result};
use crate::package::xldm::{
    OlapProofLimits, StorageProfile, Xldm140TimeGroupingContentType, generated,
    inspect as inspect_xldm, metadata, native, olap, olapproof, prove_xldm140_closure,
};

const NATIVE_MODEL_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData";
const MODEL_CONNECTION_EXTENSION_URI: &str = "{DE250136-89BD-433C-8126-D09CA5730AF9}";

use super::codec::{
    parse_document_with_base, rewrite_data_model_extension, rewrite_load_version,
    validate_definition, workbook_definition, write_data_model,
};
#[cfg(test)]
use super::model::{CalculatedTimeColumn, ModelTimeGrouping};
use super::model::{
    Definition, Model, ModelTimeGroupingContentType, ModelTimeGroupings, ModelView,
};
use super::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP_TYPE, DATA_MODEL_CONTENT_TYPE,
    DATA_MODEL_EXTENSION_URI, DATA_MODEL_PART_NAME, MAX_PAYLOAD_BYTES,
    STRICT_CONNECTIONS_RELATIONSHIP_TYPE, invalid, limit,
};

/// Load an independently owned Data Model, explicitly copying its inert payload.
/// Use [`Snapshot::load`] for a view that shares the package allocation.
pub fn load_data_model(package: &OpcPackage, workbook_name: &PackURI) -> Result<Option<Model>> {
    Ok(load_shared_model(package, workbook_name)?.map(|model| model.to_owned()))
}

fn load_shared_model(package: &OpcPackage, workbook_name: &PackURI) -> Result<Option<ModelView>> {
    let workbook = package.get_part(workbook_name)?;
    let workbook_root = parse_document_with_base(workbook.blob(), Some(workbook_name.as_str()))?;
    let (_, definition) = workbook_definition(&workbook_root)?;
    let mut parts = package
        .iter_parts()
        .filter(|part| part.content_type() == DATA_MODEL_CONTENT_TYPE);
    let part = parts.next();
    if parts.next().is_some() {
        return Err(invalid("package contains multiple Data Model parts"));
    }
    let (definition, part) = match (definition, part) {
        (Some(definition), Some(part)) => (definition, part),
        (Some(_), None) => {
            return Err(invalid(
                "workbook dataModel extension has no Data Model part",
            ));
        },
        (None, Some(_)) => {
            return Err(invalid(
                "Data Model part has no workbook dataModel extension",
            ));
        },
        (None, None) => return Ok(None),
    };
    if part.partname().as_str() != DATA_MODEL_PART_NAME {
        return Err(invalid(format!(
            "Data Model part '{}' must be '{DATA_MODEL_PART_NAME}'",
            part.partname()
        )));
    }
    // The Data Model payload is read, so it is decoded here (ADR 0030).
    let decoded = package.get_part(part.partname())?;
    let blob = decoded.blob();
    if blob.is_empty() {
        return Err(invalid("Data Model payload cannot be empty"));
    }
    if blob.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("payload bytes"));
    }
    let profile = inspect_xldm(blob)?.profile();
    if !part.rels().is_empty() {
        return Err(invalid(
            "Data Model part has forbidden outbound relationships",
        ));
    }
    validate_inbound_relationships(
        package,
        part.partname(),
        (profile == StorageProfile::Tabular150).then_some(workbook_name),
    )?;
    validate_model_time_grouping_references(&definition)?;
    validate_connections(package, workbook_name, &definition)?;
    Ok(Some(ModelView {
        definition: Arc::new(definition),
        data: decoded.blob_arc(),
    }))
}

/// Store a singleton Data Model after validating the complete mutation plan.
///
/// The outer XLDM storage profile, descriptor, package ownership, and
/// relationship closure are checked. A model-time-grouping write additionally
/// requires the neutral XLDM owner to prove the source and calculated-column
/// identity closure before the descriptor is published.
pub fn store_data_model(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &Model,
) -> Result<()> {
    if package.is_signed() || package.requires_signature_edit_policy() {
        return Err(Error::Signed);
    }
    if load_shared_model(package, workbook_name)?.is_some() {
        return Err(invalid("workbook already contains a Data Model"));
    }
    validate_model_time_grouping_identity(
        &value.payload.data,
        value.definition.model_time_groupings()?.as_ref(),
    )?;
    validate_model_contents(
        package,
        workbook_name,
        &value.definition,
        &value.payload.part_name,
        &value.payload.data,
    )?;
    let model = prepare_import(package, workbook_name, into_shared_model(value.clone())?)?;
    let mut candidate = package.clone();
    publish_data_model(&mut candidate, workbook_name, Some(&model))?;
    *package = candidate;
    Ok(())
}

fn publish_data_model(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: Option<&ModelView>,
) -> Result<()> {
    if let Some(value) = value {
        validate_model(package, workbook_name, value)?;
        if package
            .iter_parts()
            .filter(|part| part.content_type() == DATA_MODEL_CONTENT_TYPE)
            .count()
            > 1
        {
            return Err(invalid("package contains multiple Data Model parts"));
        }
        // The content type is part metadata, so the exact-name part is
        // inspected without decoding its payload (ADR 0030).
        if package.iter_parts().any(|part| {
            part.partname().as_str() == DATA_MODEL_PART_NAME
                && part.content_type() != DATA_MODEL_CONTENT_TYPE
        }) {
            return Err(invalid(format!(
                "part '{DATA_MODEL_PART_NAME}' has an unexpected content type"
            )));
        }
    }
    let removal = if value.is_none() {
        load_shared_model(package, workbook_name)?
            .map(|model| super::removal::prepare(package, workbook_name, model.definition()))
            .transpose()?
    } else {
        None
    };
    let relationship_update = if removal.is_some() {
        let id = package
            .get_part(workbook_name)?
            .rels()
            .iter()
            .find(|rel| rel.reltype() == NATIVE_MODEL_RELATIONSHIP_TYPE)
            .map(|rel| rel.r_id().to_owned());
        id.map(|id| -> Result<_> {
            let before = package.source_relationships(workbook_name)?;
            let after = before.without_relationship(&id, before.bytes().len())?;
            Ok((before, after))
        })
        .transpose()?
    } else {
        None
    };
    let workbook_source = package.source_xml_part(workbook_name)?;
    let root = parse_document_with_base(workbook_source.bytes(), Some(workbook_name.as_str()))?;
    let (core, existing) = workbook_definition(&root)?;
    let updated = match (existing.as_ref(), value) {
        (Some(before), Some(after)) if same_structure(before, &after.definition) => {
            if before.min_version_load == after.definition.min_version_load {
                None
            } else {
                Some(rewrite_load_version(
                    &workbook_source,
                    core,
                    after.definition.min_version_load,
                )?)
            }
        },
        _ => {
            let fragment = value
                .map(|value| write_data_model_fragment(core, &value.definition))
                .transpose()?;
            Some(rewrite_data_model_extension(
                &workbook_source,
                core,
                fragment.as_deref(),
            )?)
        },
    };
    let uri = PackURI::new(DATA_MODEL_PART_NAME).map_err(invalid)?;
    match value {
        // Only `PartNotFound` proves the Data Model part absent; any other
        // refusal propagates rather than reading as absence (ADR 0030).
        Some(value) => match package.get_part_mut(&uri) {
            Ok(part) => {
                if part.content_type() != DATA_MODEL_CONTENT_TYPE {
                    return Err(invalid(format!(
                        "part '{DATA_MODEL_PART_NAME}' has an unexpected content type"
                    )));
                }
                part.set_blob_shared(Arc::clone(&value.data));
            },
            Err(litchi_opc::OpcError::PartNotFound(_)) => {
                package.try_add_part(Box::new(BlobPart::new_shared(
                    uri,
                    DATA_MODEL_CONTENT_TYPE.into(),
                    Arc::clone(&value.data),
                )))?;
            },
            Err(error) => return Err(error.into()),
        },
        None => {
            let modeled = match package.get_part(&uri) {
                Ok(part) => Some(part.content_type() == DATA_MODEL_CONTENT_TYPE),
                Err(litchi_opc::OpcError::PartNotFound(_)) => None,
                Err(error) => return Err(error.into()),
            };
            match modeled {
                Some(true) => {
                    package.remove_part(&uri);
                },
                Some(false) => {
                    return Err(invalid(format!(
                        "part '{DATA_MODEL_PART_NAME}' has an unexpected content type"
                    )));
                },
                None => {},
            }
        },
    }
    if let Some(updated) = updated {
        package.try_replace_owned_xml_part(workbook_source.bytes(), updated)?;
    }
    if let Some((before, after)) = relationship_update {
        package.try_replace_relationships(&before, &after)?;
    }
    if let Some(removal) = removal {
        removal.publish(package)?;
    }
    Ok(())
}

fn write_data_model_fragment(core: &str, definition: &Definition) -> Result<Vec<u8>> {
    let descriptor = write_data_model(definition)?;
    let mut fragment = Vec::new();
    fragment.extend_from_slice(b"<x:ext xmlns:x=\"");
    escape(&mut fragment, core);
    fragment.extend_from_slice(b"\" uri=\"");
    escape(&mut fragment, DATA_MODEL_EXTENSION_URI);
    fragment.extend_from_slice(b"\">");
    fragment.extend_from_slice(&descriptor);
    fragment.extend_from_slice(b"</x:ext>");
    Ok(fragment)
}

// Normalize the detached opaque subtree's namespace and XML context closure.
// Its prefixes, values, mixed content, and comments remain lexical bytes.
fn prepare_import(
    package: &OpcPackage,
    workbook: &PackURI,
    mut model: ModelView,
) -> Result<ModelView> {
    if let Some(extension) = &mut Arc::make_mut(&mut model.definition).extension_list {
        extension.xml = super::codec::close_extension(&extension.xml)?;
        let source = package.source_xml_part(workbook)?;
        let root = parse_document_with_base(source.bytes(), Some(workbook.as_str()))?;
        let (core, _) = workbook_definition(&root)?;
        let fragment = write_data_model_fragment(core, &model.definition)?;
        let projected = rewrite_data_model_extension(&source, core, Some(&fragment))?;
        let projected_root = parse_document_with_base(projected.bytes(), Some(workbook.as_str()))?;
        model.definition = Arc::new(
            workbook_definition(&projected_root)?
                .1
                .ok_or_else(|| invalid("missing imported descriptor"))?,
        );
    }
    Ok(model)
}

fn into_shared_model(value: Model) -> Result<ModelView> {
    if value.payload.part_name != DATA_MODEL_PART_NAME {
        return Err(invalid(format!(
            "Data Model part must be '{DATA_MODEL_PART_NAME}'"
        )));
    }
    Ok(ModelView {
        definition: Arc::new(value.definition),
        data: Arc::new(value.payload.data),
    })
}

fn same_structure(left: &Definition, right: &Definition) -> bool {
    left.tables == right.tables
        && left.relationships == right.relationships
        && left.extension_list == right.extension_list
}

fn validate_model(package: &OpcPackage, workbook_name: &PackURI, value: &ModelView) -> Result<()> {
    validate_model_contents(
        package,
        workbook_name,
        &value.definition,
        value.part_name(),
        value.payload(),
    )
}

const MODEL_TIME_GROUPING_IDENTITY_FEATURE: &str = "modelTimeGroupings writes require validated inner XLDM table, column, and calculated-column identity closure";

fn model_time_grouping_identity_unsupported() -> Error {
    Error::Unsupported {
        feature: MODEL_TIME_GROUPING_IDENTITY_FEATURE,
    }
}

/// Prove every workbook time grouping against the complete neutral XLDM 140
/// closure. The proof is source-bound: it resolves the table XML name, the
/// table-local source column, every calculated column, the Date DBType of the
/// source, the integral calculated result types, and each known
/// modelTimeGrouping content type through the same inspected payload.
/// Tabular-150 payloads and incomplete/opaque closures remain readable but
/// cannot be authored through this typed operation.
fn validate_model_time_grouping_identity(
    data: &[u8],
    groupings: Option<&ModelTimeGroupings>,
) -> Result<()> {
    let Some(groupings) = groupings else {
        return Ok(());
    };
    let storage = inspect_xldm(data)?;
    if storage.profile() != StorageProfile::Xldm140 {
        return Err(model_time_grouping_identity_unsupported());
    }
    let metadata =
        metadata::inspect(&storage).map_err(|_| model_time_grouping_identity_unsupported())?;
    let native = native::inspect(&storage, &metadata.native_parse_options())
        .map_err(|_| model_time_grouping_identity_unsupported())?;
    let generated = generated::inspect_system_generated(&storage)
        .map_err(|_| model_time_grouping_identity_unsupported())?;
    let olap = olap::inspect(&storage, &metadata)
        .map_err(|_| model_time_grouping_identity_unsupported())?;
    // The table-local projection is necessary but does not prove the
    // standalone section-2.6 graph.  In particular, a grouping can appear
    // to name valid columns while a duplicate or unlinked Dimension
    // relationship remains outside that projection.  Require the same full
    // semantic proof at the XLSX authoring boundary before binding any
    // descriptor IDs.
    let olap_proof =
        olapproof::prove_xldm140_olap(&storage, &metadata, &olap, OlapProofLimits::default())
            .map_err(|_| model_time_grouping_identity_unsupported())?;
    if !olap_proof.is_complete() {
        return Err(model_time_grouping_identity_unsupported());
    }
    let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
        .map_err(|_| model_time_grouping_identity_unsupported())?;
    for grouping in &groupings.groupings {
        let calculated = grouping
            .calculated_time_columns
            .iter()
            .map(|column| {
                Ok((
                    column.column_name.as_str(),
                    column.column_id.as_str(),
                    xldm_time_grouping_content_type(&column.content_type)?,
                ))
            })
            .collect::<Result<Vec<_>>>()
            .map_err(|_| model_time_grouping_identity_unsupported())?;
        closure
            .bind_time_grouping_with_content_types(
                &grouping.table_name,
                &grouping.column_name,
                &grouping.column_id,
                &calculated,
            )
            .map_err(|_| model_time_grouping_identity_unsupported())?;
    }
    Ok(())
}

fn xldm_time_grouping_content_type(
    value: &ModelTimeGroupingContentType,
) -> Result<Xldm140TimeGroupingContentType> {
    Ok(match value {
        ModelTimeGroupingContentType::Years => Xldm140TimeGroupingContentType::Years,
        ModelTimeGroupingContentType::Quarters => Xldm140TimeGroupingContentType::Quarters,
        ModelTimeGroupingContentType::MonthsIndex => Xldm140TimeGroupingContentType::MonthsIndex,
        ModelTimeGroupingContentType::Months => Xldm140TimeGroupingContentType::Months,
        ModelTimeGroupingContentType::DaysIndex => Xldm140TimeGroupingContentType::DaysIndex,
        ModelTimeGroupingContentType::Days => Xldm140TimeGroupingContentType::Days,
        ModelTimeGroupingContentType::Hours => Xldm140TimeGroupingContentType::Hours,
        ModelTimeGroupingContentType::Minutes => Xldm140TimeGroupingContentType::Minutes,
        ModelTimeGroupingContentType::Seconds => Xldm140TimeGroupingContentType::Seconds,
        ModelTimeGroupingContentType::Other(_) => {
            return Err(model_time_grouping_identity_unsupported());
        },
    })
}

fn require_unchanged_model_time_groupings(
    before: Option<&ModelView>,
    candidate: &ModelView,
) -> Result<()> {
    let before = before
        .map(|model| model.definition.model_time_groupings())
        .transpose()?;
    let after = candidate.definition.model_time_groupings()?;
    let unchanged = match before {
        None => after.is_none(),
        Some(before) => before == after,
    };
    if !unchanged {
        return Err(Error::Unsupported {
            feature: "modelTimeGroupings writes require validated inner XLDM table, column, and calculated-column identity closure",
        });
    }
    Ok(())
}

fn validate_model_contents(
    package: &OpcPackage,
    workbook_name: &PackURI,
    definition: &Definition,
    part_name: &str,
    data: &[u8],
) -> Result<()> {
    validate_definition(definition, false)?;
    validate_model_time_grouping_references(definition)?;
    if part_name != DATA_MODEL_PART_NAME {
        return Err(invalid(format!(
            "Data Model part must be '{DATA_MODEL_PART_NAME}'"
        )));
    }
    if data.is_empty() {
        return Err(invalid("Data Model payload cannot be empty"));
    }
    if data.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("payload bytes"));
    }
    let profile = inspect_xldm(data)?.profile();
    validate_inbound_relationships(
        package,
        &PackURI::new(part_name).map_err(invalid)?,
        (profile == StorageProfile::Tabular150).then_some(workbook_name),
    )?;
    validate_connections(package, workbook_name, definition)
}

fn validate_model_time_grouping_references(definition: &Definition) -> Result<()> {
    let Some(groupings) = definition.model_time_groupings()? else {
        return Ok(());
    };
    for grouping in groupings.groupings {
        if !definition
            .tables
            .iter()
            .any(|table| table.name.eq_ignore_ascii_case(&grouping.table_name))
        {
            return Err(invalid(format!(
                "modelTimeGrouping references unknown table '{}'",
                grouping.table_name
            )));
        }
    }
    Ok(())
}

fn validate_connections(
    package: &OpcPackage,
    workbook_name: &PackURI,
    definition: &Definition,
) -> Result<()> {
    if definition.tables.is_empty() {
        return Ok(());
    }
    let workbook = package.get_part(workbook_name)?;
    let mut relationships = workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            CONNECTIONS_RELATIONSHIP_TYPE | STRICT_CONNECTIONS_RELATIONSHIP_TYPE
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("Data Model tables require a workbook Connections part"))?;
    if relationships.next().is_some() {
        return Err(invalid("workbook has multiple Connections relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("Connections relationship cannot be external"));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(invalid(
            "Connections relationship cannot have a target query or fragment",
        ));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(invalid(format!(
            "Connections part '{target}' has content type '{}', expected '{CONNECTIONS_CONTENT_TYPE}'",
            part.content_type()
        )));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Connections part has forbidden outbound relationships",
        ));
    }
    let connections = crate::connections::Connections::parse(part.blob())
        .map_err(|error| invalid(format!("invalid Connections part: {error}")))?;
    validate_model_connection_extensions(part.blob())?;
    let mut names = HashMap::<String, usize>::new();
    for name in connections
        .connections
        .iter()
        .filter_map(|connection| connection.name.as_deref())
    {
        *names.entry(normalized_connection_name(name)?).or_default() += 1;
    }
    for table in &definition.tables {
        if names
            .get(&normalized_connection_name(&table.connection)?)
            .copied()
            != Some(1)
        {
            return Err(invalid(format!(
                "Data Model table '{}' requires exactly one workbook connection named '{}'",
                table.name, table.connection
            )));
        }
    }
    Ok(())
}

fn normalized_connection_name(value: &str) -> Result<String> {
    crate::raw::strings::decode_spreadsheet_text(value).map(|value| value.to_lowercase())
}

fn connection_attribute<'a>(node: &'a super::codec::Node, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace.is_empty() && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}

fn connection_boolean(value: &str) -> Result<bool> {
    match value.trim_matches([' ', '\t', '\r', '\n']) {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!(
            "invalid model connection boolean '{value}'"
        ))),
    }
}

fn validate_model_connection_extensions(xml: &[u8]) -> Result<()> {
    let root = parse_connections_document(xml)?;
    for connection in root.children.iter().filter(|node| {
        (node.namespace == super::SML || node.namespace == super::STRICT_SML)
            && node.name == "connection"
    }) {
        let connection_type = connection_attribute(connection, "type")
            .map(|value| {
                value
                    .parse::<u32>()
                    .map_err(|_| invalid("invalid workbook connection type"))
            })
            .transpose()?;
        for ext_list in connection.children.iter().filter(|node| {
            (node.namespace == super::SML || node.namespace == super::STRICT_SML)
                && node.name == "extLst"
        }) {
            for extension in ext_list.children.iter().filter(|node| {
                (node.namespace == super::SML || node.namespace == super::STRICT_SML)
                    && node.name == "ext"
                    && connection_attribute(node, "uri") == Some(MODEL_CONNECTION_EXTENSION_URI)
            }) {
                for model_connection in extension
                    .children
                    .iter()
                    .filter(|node| node.namespace == super::X15 && node.name == "connection")
                {
                    let is_model = model_connection
                        .attributes
                        .iter()
                        .find(|attribute| {
                            attribute.namespace.is_empty() && attribute.name == "model"
                        })
                        .map(|attribute| connection_boolean(&attribute.value))
                        .transpose()?
                        .unwrap_or(false);
                    if !is_model {
                        continue;
                    }
                    // MS-XLSX 2.4.21 and 2.6.91 make the outer connection a
                    // type-5 model connection and require its ST_Xstring id
                    // to have zero decoded characters.
                    if connection_type != Some(5) {
                        return Err(invalid("model workbook connection must have type 5"));
                    }
                    let id = connection_attribute(model_connection, "id")
                        .ok_or_else(|| invalid("model workbook connection requires an id"))?;
                    if !crate::raw::strings::decode_spreadsheet_text(id)?.is_empty() {
                        return Err(invalid("model workbook connection id must be empty"));
                    }
                }
            }
        }
    }
    Ok(())
}

// Connections::parse permits processing instructions because they are legal
// inert XML in this part. The Data Model codec deliberately rejects them, so
// remove only those events before reusing its bounded namespace tree for the
// model-connection identity check.
pub(super) fn parse_connections_document(xml: &[u8]) -> Result<super::codec::Node> {
    let mut reader = Reader::from_reader(xml);
    let mut filtered: Option<Vec<u8>> = None;
    let mut cursor = 0;
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::PI(_)) => {
                let end = reader.buffer_position() as usize;
                if start < cursor || end < start || end > xml.len() {
                    return Err(invalid("invalid Connections processing-instruction span"));
                }
                if filtered.is_none() {
                    let mut bytes = Vec::new();
                    bytes
                        .try_reserve_exact(xml.len() - (end - start))
                        .map_err(|_| {
                            invalid("Connections XML processing-instruction allocation failed")
                        })?;
                    filtered = Some(bytes);
                }
                if let Some(bytes) = filtered.as_mut() {
                    bytes.extend_from_slice(&xml[cursor..start]);
                }
                cursor = end;
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(invalid(format!("invalid Connections XML: {error}"))),
        }
    }
    match filtered {
        Some(mut bytes) => {
            bytes.extend_from_slice(&xml[cursor..]);
            super::codec::parse_document(&bytes)
        },
        None => super::codec::parse_document(xml),
    }
}

// The native tabular profile retains the optional workbook powerPivotData edge
// observed in the fixture. Other sources/types, duplicate edges, external
// targets, and URI suffixes cannot enter that compatibility rule.
fn validate_inbound_relationships(
    package: &OpcPackage,
    target: &PackURI,
    native_workbook: Option<&PackURI>,
) -> Result<()> {
    for relationship in package.rels().iter() {
        if relationship.reltype() == NATIVE_MODEL_RELATIONSHIP_TYPE
            || !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
        {
            return Err(invalid(
                "package relationship targets the relationship-free Data Model part",
            ));
        }
    }
    let mut native_edges = 0usize;
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship.reltype() == NATIVE_MODEL_RELATIONSHIP_TYPE {
                if native_workbook != Some(source.partname())
                    || relationship.is_external()
                    || !relationship.target_partname()?.is_equivalent_to(target)
                    || relationship.target_query().is_some()
                    || relationship.target_fragment().is_some()
                {
                    return Err(invalid("invalid native workbook Data Model relationship"));
                }
                native_edges += 1;
                if native_edges > 1 {
                    return Err(invalid(
                        "duplicate native workbook Data Model relationships",
                    ));
                }
                continue;
            }
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                return Err(invalid(format!(
                    "part '{}' has a relationship to the relationship-free Data Model part",
                    source.partname()
                )));
            }
        }
    }
    Ok(())
}

fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipState {
    source: PackURI,
    id: String,
    reltype: String,
    target: String,
    mode: TargetMode,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourcePart {
    name: PackURI,
    content_type: Box<str>,
    blob: Arc<Vec<u8>>,
    relationships: litchi_opc::OwnedRelationships,
    xml: Option<litchi_opc::OwnedXmlPart>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourceState {
    content_types: litchi_opc::OwnedContentTypes,
    workbook_blob: Arc<Vec<u8>>,
    workbook_proof: litchi_opc::OwnedXmlPart,
    workbook_relationships: litchi_opc::OwnedRelationships,
    model_part: Option<SourcePart>,
    model_incoming_relationships: Arc<[RelationshipState]>,
    connections_part: Option<SourcePart>,
}

/// An immutable, source-bound Data Model snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    workbook_name: PackURI,
    model: Option<ModelView>,
    source: Arc<SourceState>,
}

impl Snapshot {
    /// Read the workbook's optional Data Model using the canonical workbook
    /// relationship, without exposing the opaque payload internals.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let workbook_name = package.main_document_part()?.partname().clone();
        Self::load_for(package, workbook_name)
    }

    fn load_for(package: &OpcPackage, workbook_name: PackURI) -> Result<Self> {
        let model = load_shared_model(package, &workbook_name)?;
        let source = Arc::new(capture_source(package, &workbook_name, model.as_ref())?);
        Ok(Self {
            workbook_name,
            model,
            source,
        })
    }

    /// The typed Data Model descriptor plus its inert payload, when present.
    #[must_use]
    pub fn model(&self) -> Option<&ModelView> {
        self.model.as_ref()
    }

    /// Whether this workbook has no Data Model graph.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.model.is_none()
    }

    fn same_source(&self, other: &Self) -> bool {
        self.workbook_name == other.workbook_name && self.source == other.source
    }

    fn same_semantics(&self, model: Option<&ModelView>) -> bool {
        self.model.as_ref() == model
    }
}

/// A clone-staged source-bound Data Model transaction.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Option<ModelView>,
}

impl<'a> Transaction<'a> {
    /// Start a transaction from the package's current Data Model graph.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        Ok(Self {
            draft: before.model.clone(),
            target,
            before,
        })
    }

    /// Read the source snapshot captured at transaction start.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Read the currently staged model, if any.
    #[must_use]
    pub fn model(&self) -> Option<&ModelView> {
        self.draft.as_ref()
    }

    /// Replace the complete model after validating its descriptor, payload,
    /// and workbook connection dependencies.
    ///
    /// Structural table, relationship, or opaque-extension replacement is
    /// refused for an existing source-bound model. The outer descriptor and
    /// XLDM storage profile cannot prove that a changed opaque payload keeps
    /// the same inner table, relationship, column, and time-group identity
    /// closure. Creation into an empty workbook remains supported through the
    /// outer validation performed by [`store_data_model`].
    /// Imported opaque extensions acquire explicit namespace context for the
    /// destination workbook; inspect [`Self::model`] for the staged definition.
    pub fn set(&mut self, model: Model) -> Result<bool> {
        let model = into_shared_model(model)?;
        if self.draft.as_ref() == Some(&model) {
            return Ok(false);
        }
        self.check_structure(&model)?;
        let model = if model.definition.extension_list.is_some()
            && !self
                .before
                .model
                .as_ref()
                .is_some_and(|before| same_structure(&before.definition, &model.definition))
        {
            prepare_import(self.target, &self.before.workbook_name, model)?
        } else {
            model
        };
        self.set_shared(model)
    }

    fn set_shared(&mut self, model: ModelView) -> Result<bool> {
        if self.draft.as_ref() == Some(&model) {
            return Ok(false);
        }
        self.check_structure(&model)?;
        require_unchanged_model_time_groupings(self.before.model.as_ref(), &model)?;
        // A payload-only replacement can invalidate the identities and
        // derivation closure even when the outer descriptor is byte-for-byte
        // unchanged. Reprove the retained extension against the candidate
        // payload before staging it, so set/edit_definition cannot leave stale
        // modelTimeGroupings attached to a different XLDM source.
        validate_model_time_grouping_identity(
            model.payload(),
            model.definition.model_time_groupings()?.as_ref(),
        )?;
        validate_model(self.target, &self.before.workbook_name, &model)?;
        self.draft = Some(model);
        Ok(true)
    }

    fn check_structure(&self, model: &ModelView) -> Result<()> {
        // A complete replacement must not disguise an unverified inner XLDM
        // structural change by supplying different opaque payload bytes. Check
        // the original as well, so remove/set cannot bypass this refusal.
        for previous in [self.before.model.as_ref(), self.draft.as_ref()]
            .into_iter()
            .flatten()
        {
            if !same_structure(&previous.definition, &model.definition) {
                return Err(Error::Unsupported {
                    feature: "structural Data Model edits require validated inner XLDM identity closure",
                });
            }
        }
        Ok(())
    }

    /// Edit the load-version descriptor while retaining the opaque payload.
    /// Structural table, relationship, and extension changes require a complete
    /// replacement model with matching XLDM bytes through [`Self::set`].
    pub fn edit_definition(
        &mut self,
        edit: impl FnOnce(&mut Definition) -> Result<()>,
    ) -> Result<bool> {
        let current = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid("workbook has no Data Model to edit"))?;
        let mut definition = current.definition.as_ref().clone();
        edit(&mut definition)?;
        let candidate = ModelView {
            definition: Arc::new(definition),
            data: Arc::clone(&current.data),
        };
        self.set_shared(candidate)
    }

    /// Edit only the typed `modelTimeGroupings` extension while retaining the
    /// source-bound XLDM payload and every unrelated extension byte. The
    /// neutral XLDM closure proves table, source-column, and calculated-column
    /// identity before the edit is staged.
    pub fn edit_model_time_groupings(&mut self, value: Option<ModelTimeGroupings>) -> Result<bool> {
        let current = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid("workbook has no Data Model to edit"))?;
        if current.definition.model_time_groupings()? == value {
            return Ok(false);
        }
        let mut definition = current.definition.as_ref().clone();
        definition.set_model_time_groupings(value)?;
        let candidate = ModelView {
            definition: Arc::new(definition),
            data: Arc::clone(&current.data),
        };
        validate_model_time_grouping_identity(
            candidate.payload(),
            candidate.definition.model_time_groupings()?.as_ref(),
        )?;
        validate_model(self.target, &self.before.workbook_name, &candidate)?;
        self.draft = Some(candidate);
        Ok(true)
    }

    /// Remove the staged Data Model graph.
    pub fn remove(&mut self) -> Result<bool> {
        if self.draft.is_none() {
            return Ok(false);
        }
        self.draft = None;
        Ok(true)
    }

    /// Whether the typed model or payload differs from the source snapshot.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.before.same_semantics(self.draft.as_ref())
    }

    /// Validate the source closure and atomically publish the staged model.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            return Ok(Commit::new(
                self.before.clone(),
                Patch::new(self.before.clone(), self.before.clone()),
                false,
            ));
        }
        if self.target.is_signed() || self.target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load(self.target)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "Data Model source closure".into(),
            });
        }
        let mut candidate = self.target.clone();
        publish_data_model(
            &mut candidate,
            &self.before.workbook_name,
            self.draft.as_ref(),
        )?;
        let snapshot = Snapshot::load(&candidate)?;
        if !snapshot.same_semantics(self.draft.as_ref()) {
            return Err(invalid("Data Model publication changed staged semantics"));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}

/// An exact, source-checked Data Model replacement.
#[derive(Debug, Clone)]
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

    /// Whether this patch is a byte-preserving no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return an exact source-bound inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply the patch atomically, restoring retained source bytes for the
    /// workbook descriptor and opaque payload.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "Data Model source closure".into(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if target.is_signed() || target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        if self.after.model.is_none() {
            if let Some(model) = &self.before.model {
                super::removal::prepare(target, &self.before.workbook_name, model.definition())?;
            }
        }
        if self.before.source.connections_part.is_none() {
            if let Some(restored) = &self.after.source.connections_part {
                // Only `PartNotFound` proves the restored owner absent; any
                // other refusal propagates (ADR 0030).
                match target.get_part(&restored.name) {
                    Ok(_) => {
                        return Err(Error::PatchConflict {
                            part: "Data Model restored Connections owner".into(),
                        });
                    },
                    Err(litchi_opc::OpcError::PartNotFound(_)) => {},
                    Err(error) => return Err(error.into()),
                }
                if super::removal::has_foreign_incoming(target, &restored.name, None)? {
                    return Err(Error::PatchConflict {
                        part: "Data Model restored Connections references".into(),
                    });
                }
            }
        }
        let mut candidate = target.clone();
        publish_snapshot(&mut candidate, &self.before, &self.after)?;
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after) {
            return Err(invalid("Data Model patch publication changed source bytes"));
        }
        *target = candidate;
        Ok(())
    }
}

/// Successful Data Model transaction publication.
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

    /// Whether the model graph changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting source-bound snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch produced by the transaction.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }
}

fn capture_source(
    package: &OpcPackage,
    workbook_name: &PackURI,
    _model: Option<&ModelView>,
) -> Result<SourceState> {
    let workbook = package.get_part(workbook_name)?;
    let model_uri = PackURI::new(DATA_MODEL_PART_NAME).map_err(invalid)?;
    // Only `PartNotFound` proves the model part absent; any other refusal
    // propagates rather than capturing a source without it (ADR 0030).
    let model_part = match package.get_part(&model_uri) {
        Ok(part) => Some(source_part(package, part)?),
        Err(litchi_opc::OpcError::PartNotFound(_)) => None,
        Err(error) => return Err(error.into()),
    };
    let model_incoming_relationships =
        Arc::from(incoming_model_relationships(package, &model_uri)?);
    let mut connections = workbook.rels().iter().filter(|rel| {
        matches!(
            rel.reltype(),
            CONNECTIONS_RELATIONSHIP_TYPE | STRICT_CONNECTIONS_RELATIONSHIP_TYPE
        )
    });
    let connection = connections.next();
    if connections.next().is_some() {
        return Err(invalid("workbook has multiple Connections relationships"));
    }
    let connections_part = connection
        .map(|rel| -> Result<_> {
            if rel.is_external() {
                return Err(invalid("external Connections relationship"));
            }
            source_part(package, package.get_part(&rel.target_partname()?)?)
        })
        .transpose()?;
    Ok(SourceState {
        content_types: package.source_content_types()?,
        workbook_blob: workbook.blob_arc(),
        workbook_proof: package.source_xml_part(workbook_name)?,
        workbook_relationships: package.source_relationships(workbook_name)?,
        model_part,
        model_incoming_relationships,
        connections_part,
    })
}

fn source_part(package: &OpcPackage, part: &dyn Part) -> Result<SourcePart> {
    Ok(SourcePart {
        name: part.partname().clone(),
        content_type: part.content_type().into(),
        blob: part.blob_arc(),
        relationships: package.source_relationships(part.partname())?,
        xml: (part.content_type() == CONNECTIONS_CONTENT_TYPE)
            .then(|| package.source_xml_part(part.partname()))
            .transpose()?,
    })
}

fn incoming_model_relationships(
    package: &OpcPackage,
    target: &PackURI,
) -> Result<Vec<RelationshipState>> {
    let mut values = Vec::new();
    let package_source = PackURI::new("/").map_err(invalid)?;
    for relationship in package.rels().iter() {
        if relationship.is_external() {
            continue;
        }
        if relationship.target_partname()?.is_equivalent_to(target) {
            values.push(RelationshipState {
                source: package_source.clone(),
                id: relationship.r_id().to_owned(),
                reltype: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                mode: relationship.target_mode(),
            });
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            if relationship.target_partname()?.is_equivalent_to(target) {
                values.push(RelationshipState {
                    source: source.partname().clone(),
                    id: relationship.r_id().to_owned(),
                    reltype: relationship.reltype().to_owned(),
                    target: relationship.target_ref().to_owned(),
                    mode: relationship.target_mode(),
                });
            }
        }
    }
    values.sort_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.id.cmp(&right.id))
    });
    Ok(values)
}

fn publish_snapshot(
    package: &mut OpcPackage,
    before: &Snapshot,
    snapshot: &Snapshot,
) -> Result<()> {
    let current = package.get_part(&snapshot.workbook_name)?.blob_arc();
    package.try_replace_owned_xml_part(&current, snapshot.source.workbook_proof.clone())?;
    let current_relationships = package.source_relationships(&snapshot.workbook_name)?;
    package.try_replace_relationships(
        &current_relationships,
        &snapshot.source.workbook_relationships,
    )?;
    let model_uri = PackURI::new(DATA_MODEL_PART_NAME).map_err(invalid)?;
    match snapshot.source.model_part.as_ref() {
        Some(part) => publish_source_part(package, part)?,
        None => match package.get_part(&model_uri) {
            Ok(_) => {
                package.remove_part(&model_uri);
            },
            Err(litchi_opc::OpcError::PartNotFound(_)) => {},
            Err(error) => return Err(error.into()),
        },
    }
    match (
        &before.source.connections_part,
        &snapshot.source.connections_part,
    ) {
        (_, Some(part)) => publish_source_part(package, part)?,
        (Some(part), None) => {
            if !package.remove_part(&part.name) {
                return Err(invalid(
                    "missing Connections owner during patch publication",
                ));
            }
        },
        (None, None) => {},
    }
    let current_content_types = package.source_content_types()?;
    package.try_replace_content_types(
        current_content_types.bytes(),
        &snapshot.source.content_types,
    )?;
    Ok(())
}

fn publish_source_part(package: &mut OpcPackage, part: &SourcePart) -> Result<()> {
    // Only `PartNotFound` proves the part absent; any other refusal
    // propagates rather than taking the add path (ADR 0030).
    let existing = match package.get_part(&part.name) {
        Ok(existing) => Some(existing),
        Err(litchi_opc::OpcError::PartNotFound(_)) => None,
        Err(error) => return Err(error.into()),
    };
    if let Some(existing) = existing {
        if existing.content_type() != part.content_type.as_ref() {
            return Err(invalid("Data Model source part content type changed"));
        }
        if let Some(xml) = &part.xml {
            let expected = existing.blob_arc();
            package.try_replace_owned_xml_part(&expected, xml.clone())?;
        } else {
            package
                .get_part_mut(&part.name)?
                .set_blob_shared(Arc::clone(&part.blob));
        }
    } else if let Some(xml) = &part.xml {
        package.try_add_owned_xml_part(xml.clone())?;
    } else {
        package.try_add_part(Box::new(BlobPart::new_shared(
            part.name.clone(),
            part.content_type.to_string(),
            Arc::clone(&part.blob),
        )))?;
    }
    let current = package.source_relationships(&part.name)?;
    package.try_replace_relationships(&current, &part.relationships)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::codec::{
        parse_data_model, parse_document, parse_document_with_base, rewrite_data_model_extension,
        workbook_definition,
    };
    use super::*;
    use crate::package::xldm_test_support::test_xldm_bytes;
    use crate::workbook::data_model::{
        MAX_XML_BYTES, MODEL_TIME_GROUPINGS_EXTENSION_URI, Payload, SML, X15,
    };
    use litchi_opc::Part;

    fn definition() -> Definition {
        Definition {
            min_version_load: 7,
            tables: vec![
                super::super::model::Table {
                    id: "t-sales".into(),
                    name: "Sales".into(),
                    connection: "ModelConnection".into(),
                },
                super::super::model::Table {
                    id: "t-date".into(),
                    name: "Date".into(),
                    connection: "ModelConnection".into(),
                },
            ],
            relationships: vec![super::super::model::Relationship {
                from_table: "Sales".into(),
                from_column: "DateKey".into(),
                to_table: "Date".into(),
                to_column: "DateKey".into(),
            }],
            extension_list: Some(super::super::model::OpaqueXml {
                xml: format!(
                    r#"<x15:extLst xmlns:x15="{X15}"><x15:ext uri="urn:test"><v:opaque xmlns:v="urn:vendor"/></x15:ext></x15:extLst>"#
                )
                .into_bytes(),
            }),
        }
    }

    fn fixture_package() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let mut part = BlobPart::new(
            workbook.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            format!(r#"<workbook xmlns="{SML}"><sheets/></workbook>"#).into_bytes(),
        );
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        part.rels_mut().add_relationship(
            CONNECTIONS_RELATIONSHIP_TYPE.into(),
            "connections.xml".into(),
            "rIdConnections".into(),
            false,
        );
        package.add_part(Box::new(part));
        package.add_part(Box::new(BlobPart::new(
            connections,
            CONNECTIONS_CONTENT_TYPE.into(),
            format!(r#"<connections xmlns="{SML}"><connection id="1" name="ModelConnection" refreshedVersion="7"/></connections>"#).into_bytes(),
        )));
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "xl/workbook.xml".into(),
            "rIdWorkbook".into(),
            false,
        );
        (package, workbook)
    }

    // A synthetic producer ZIP, deliberately bypassing the authored compact
    // writer so that tests exercise genuine ingress provenance for spaced XML.
    fn as_source(package: &OpcPackage) -> OpcPackage {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        let mut parts = package
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
            .collect::<Vec<_>>();
        parts.sort_by(|a, b| a.partname().as_str().cmp(b.partname().as_str()));
        let mut types = String::from(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        );
        for part in &parts {
            types.push_str(&format!(
                r#"<Override PartName="{}" ContentType="{}"/>"#,
                part.partname(),
                part.content_type()
            ));
            writer
                .write_stored(
                    part.partname().as_str().trim_start_matches('/'),
                    part.blob(),
                )
                .unwrap();
            if !part.rels().is_empty() {
                writer
                    .write_stored(
                        part.partname()
                            .rels_uri()
                            .unwrap()
                            .as_str()
                            .trim_start_matches('/'),
                        part.rels().to_xml().as_bytes(),
                    )
                    .unwrap();
            }
        }
        types.push_str("</Types>");
        writer
            .write_stored("[Content_Types].xml", types.as_bytes())
            .unwrap();
        writer
            .write_stored("_rels/.rels", package.rels().to_xml().as_bytes())
            .unwrap();
        OpcPackage::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap()
    }

    fn model() -> Model {
        Model {
            definition: definition(),
            payload: Payload {
                part_name: DATA_MODEL_PART_NAME.into(),
                data: test_xldm_bytes(),
            },
        }
    }

    fn install_model_time_groupings(
        package: &mut OpcPackage,
        workbook: &PackURI,
        value: ModelTimeGroupings,
    ) {
        let mut current = load_data_model(package, workbook).unwrap().unwrap();
        current
            .definition
            .set_model_time_groupings(Some(value))
            .unwrap();
        let (source, updated) = {
            let source = package.source_xml_part(workbook).unwrap();
            let source_bytes = source.bytes().to_vec();
            let root = parse_document_with_base(&source_bytes, Some(workbook.as_str())).unwrap();
            let (core, _) = workbook_definition(&root).unwrap();
            let fragment = write_data_model_fragment(core, &current.definition).unwrap();
            let updated = rewrite_data_model_extension(&source, core, Some(&fragment)).unwrap();
            (source_bytes, updated)
        };
        package
            .try_replace_owned_xml_part(&source, updated)
            .unwrap();
    }

    fn replace_connections_target(package: &mut OpcPackage, workbook: &PackURI, target: &str) {
        let relationships = package.get_part_mut(workbook).unwrap().rels_mut();
        relationships.remove("rIdConnections");
        relationships.add_relationship(
            CONNECTIONS_RELATIONSHIP_TYPE.into(),
            target.into(),
            "rIdConnections".into(),
            false,
        );
    }

    fn remove_low_level_signature(package: &mut OpcPackage) {
        let relationship_id = package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());
        package.rels_mut().remove(&relationship_id);
        assert!(!package.is_signed());
        assert!(package.requires_signature_edit_policy());
    }

    fn payload_with_equivalent_object_id_case() -> Vec<u8> {
        let mut payload = test_xldm_bytes();
        let needle = "<LastWriteTime>0</LastWriteTime>"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        let offset = payload
            .windows(needle.len())
            .position(|window| window == needle)
            .expect("test XLDM directory timestamp");
        let value_offset = offset + "<LastWriteTime>".encode_utf16().count() * 2;
        payload[value_offset] = b'1';
        assert!(litchi_xldm::inspect(&payload).is_ok());
        payload
    }

    #[test]
    fn direct_model_store_rejects_unowned_incoming_edges_before_publication() {
        for reltype in ["urn:unowned", NATIVE_MODEL_RELATIONSHIP_TYPE] {
            let (mut package, workbook) = fixture_package();
            package
                .get_part_mut(&workbook)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    reltype.into(),
                    "model/item.data".into(),
                    "rIdUnowned".into(),
                    false,
                );
            let before = Snapshot::load(&package).unwrap();
            assert!(store_data_model(&mut package, &workbook, &model()).is_err());
            assert!(Snapshot::load(&package).unwrap().same_source(&before));
            assert!(
                package
                    .get_part(&PackURI::new(DATA_MODEL_PART_NAME).unwrap())
                    .is_err()
            );
        }
    }

    #[test]
    fn fresh_model_time_grouping_store_refuses_without_inner_identity_proof() {
        let (mut package, workbook) = fixture_package();
        let mut value = model();
        value
            .definition
            .set_model_time_groupings(Some(ModelTimeGroupings {
                groupings: vec![ModelTimeGrouping {
                    table_name: "Sales".into(),
                    column_name: "OrderDate".into(),
                    column_id: "date-1".into(),
                    calculated_time_columns: vec![CalculatedTimeColumn {
                        column_name: "Year".into(),
                        column_id: "year-1".into(),
                        content_type: ModelTimeGroupingContentType::Years,
                        is_selected: true,
                    }],
                }],
            }))
            .unwrap();
        assert!(matches!(
            store_data_model(&mut package, &workbook, &value),
            Err(Error::Unsupported { feature })
                if feature.contains("inner XLDM table, column, and calculated-column identity closure")
        ));
        assert!(load_data_model(&package, &workbook).unwrap().is_none());
    }

    #[test]
    fn transaction_import_preserves_opaque_descriptor_markup() {
        let (mut package, workbook) = fixture_package();
        package.get_part_mut(&workbook).unwrap().set_blob(
            format!(
                r#"<w:workbook xmlns:w="{SML}" xmlns:q="urn:destination"><w:sheets/></w:workbook>"#
            )
            .into_bytes(),
        );
        let mut expected = model();
        let content = r#"<!--before--><v:opaque xmlns:v="urn:vendor" type="q:Source">left<v:child/>right<!--inside--></v:opaque><unqualified/>"#;
        expected.definition.extension_list.as_mut().unwrap().xml =
            format!(r#"<x15:extLst xmlns:x15="{X15}" xmlns:q="urn:source">{content}</x15:extLst>"#)
                .into_bytes();
        let payload_pointer = expected.payload.data.as_ptr();
        let before = Snapshot::load(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.set(expected).unwrap();
        assert_eq!(
            transaction.model().unwrap().payload().as_ptr(),
            payload_pointer
        );
        let commit = transaction.commit().unwrap();
        let after = commit.snapshot().clone();
        let patch = commit.patch().clone();
        let xml = String::from_utf8_lossy(
            &after
                .model()
                .unwrap()
                .definition()
                .extension_list
                .as_ref()
                .unwrap()
                .xml,
        );
        assert!(xml.contains(content));
        assert!(xml.contains(r#"xmlns:q="urn:source""#));
        assert!(xml.contains(r#"xmlns="""#));
        let bytes = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(Snapshot::load(&reopened).unwrap().same_source(&after));
        patch.inverse().apply(&mut reopened).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let restored = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(Snapshot::load(&restored).unwrap().same_source(&before));
    }

    #[test]
    fn opaque_descriptor_retains_inherited_bindings_and_mixed_content() {
        let content = r#"<!--before--><vendor type='q:OnlyInValue'>left<a:child/>middle<![CDATA[<raw>]]><?vendor keep?>right<!--after--></vendor><a:scope xmlns:q='urn:shadow' type='q:Local'/>"#;
        let xml = format!(
            r#"<m:dataModel xmlns:m='{X15}' xmlns='urn:default' xmlns:a='urn:a&amp;b' xmlns:q='urn:value'><m:extLst marker='q:Tag'>{content}</m:extLst></m:dataModel>"#
        );
        let parsed = parse_data_model(xml.as_bytes()).unwrap();
        let extension = String::from_utf8_lossy(&parsed.extension_list.as_ref().unwrap().xml);
        assert!(extension.contains(content));
        assert!(extension.contains("marker='q:Tag'"));
        for declaration in [
            r#"xmlns="urn:default""#,
            r#"xmlns:a="urn:a&amp;b""#,
            r#"xmlns:q="urn:value""#,
        ] {
            assert!(extension.contains(declaration), "{extension}");
        }
        let detached = super::super::codec::close_extension(extension.as_bytes()).unwrap();
        assert_eq!(detached, extension.as_bytes());
        let serialized = write_data_model(&parsed).unwrap();
        assert!(String::from_utf8_lossy(&serialized).contains(content));
        assert!(parse_data_model(&serialized).is_ok());
        assert!(
            parse_data_model(
                format!(r#"<m:dataModel xmlns:m='{X15}'><![CDATA[bad]]></m:dataModel>"#).as_bytes()
            )
            .is_err()
        );
        assert!(
            parse_data_model(
                format!(r#"<!DOCTYPE dataModel><m:dataModel xmlns:m='{X15}'/>"#).as_bytes()
            )
            .is_err()
        );
    }

    #[test]
    fn source_opaque_content_survives_version_edit_and_inverse_save() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let xml = format!(
            r#"<w:workbook xmlns:w='{SML}' xmlns:m='{X15}' xmlns:q='urn:context'><w:sheets/><w:extLst><w:ext uri='{DATA_MODEL_EXTENSION_URI}'><m:dataModel minVersionLoad = '7'><m:extLst><!--start--><foreign type='q:Value'>left<![CDATA[<raw>]]><?vendor keep?>right<!--end--></foreign></m:extLst></m:dataModel></w:ext></w:extLst></w:workbook>"#
        );
        package
            .get_part_mut(&workbook)
            .unwrap()
            .set_blob(xml.as_bytes().to_vec());
        let mut package = as_source(&package);
        let before = Snapshot::load(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        let mut owned = transaction.model().unwrap().to_owned();
        assert!(!transaction.set(owned.clone()).unwrap());
        owned.definition.min_version_load = 8;
        transaction.set(owned).unwrap();
        let commit = transaction.commit().unwrap();
        let patch = commit.patch().clone();
        let bytes = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            reopened.get_part(&workbook).unwrap().blob(),
            xml.replace("= '7'", "= '8'").as_bytes()
        );
        patch.inverse().apply(&mut reopened).unwrap();
        let bytes = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let restored = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(Snapshot::load(&restored).unwrap().same_source(&before));
        assert_eq!(restored.get_part(&workbook).unwrap().blob(), xml.as_bytes());
    }

    #[test]
    fn opaque_namespace_capture_bounds_resolver_shadowing_work() {
        for (count, allowed) in [(4093, true), (4094, false)] {
            let mut xml = format!("<root xmlns:m='{X15}'>");
            let mut levels = 0;
            for start in (0..count).step_by(128) {
                xml.push_str("<nested");
                for index in start..(start + 128).min(count) {
                    xml.push_str(&format!(" xmlns:p{index}='urn:{index}'"));
                }
                xml.push('>');
                levels += 1;
            }
            xml.push_str("<m:dataModel><m:extLst/></m:dataModel>");
            for _ in 0..levels {
                xml.push_str("</nested>");
            }
            xml.push_str("</root>");
            let result = parse_document(xml.as_bytes());
            if allowed {
                assert!(result.is_ok());
            } else {
                assert!(
                    matches!(result, Err(Error::Invalid(message)) if message.contains("namespace count"))
                );
            }
        }
    }

    #[test]
    fn typed_descriptor_round_trip() {
        let expected = definition();
        let xml = write_data_model(&expected).unwrap();
        let actual = parse_data_model(&xml).unwrap();
        assert_eq!(actual.min_version_load, expected.min_version_load);
        assert_eq!(actual.tables, expected.tables);
        assert_eq!(actual.relationships, expected.relationships);
        assert!(String::from_utf8_lossy(&actual.extension_list.unwrap().xml).contains("opaque"));
    }

    #[test]
    fn package_round_trip_preserves_inert_payload_and_inline_metadata() {
        let (mut package, workbook) = fixture_package();
        let expected = model();
        store_data_model(&mut package, &workbook, &expected).unwrap();
        let actual = load_data_model(&package, &workbook).unwrap().unwrap();
        assert_eq!(
            actual.definition.min_version_load,
            expected.definition.min_version_load
        );
        assert_eq!(actual.definition.tables, expected.definition.tables);
        assert_eq!(
            actual.definition.relationships,
            expected.definition.relationships
        );
        assert!(
            String::from_utf8_lossy(&actual.definition.extension_list.as_ref().unwrap().xml)
                .contains("opaque")
        );
        assert_eq!(actual.payload, expected.payload);
    }

    #[test]
    fn inserts_into_existing_empty_extension_list() {
        let (mut package, workbook) = fixture_package();
        package.get_part_mut(&workbook).unwrap().set_blob(
            format!(r#"<workbook xmlns="{SML}"><sheets/><extLst /></workbook>"#).into_bytes(),
        );
        let mut package = as_source(&package);
        store_data_model(&mut package, &workbook, &model()).unwrap();
        assert!(load_data_model(&package, &workbook).unwrap().is_some());
    }

    #[test]
    fn rejects_hostile_xml_schema_and_bounds() {
        for xml in [
            format!(r#"<!DOCTYPE x><x15:dataModel xmlns:x15="{X15}"/>"#),
            format!(r#"<?bad x?><x15:dataModel xmlns:x15="{X15}"/>"#),
            format!(r#"<x15:dataModel xmlns:x15="{X15}" minVersionLoad="4"/>"#),
            format!(r#"<x15:dataModel xmlns:x15="{X15}"><x15:modelTables/></x15:dataModel>"#),
            format!(
                r#"<x15:dataModel xmlns:x15="{X15}"><x15:modelRelationships><x15:modelRelationship fromTable="Missing" fromColumn="a" toTable="Missing" toColumn="b"/></x15:modelRelationships></x15:dataModel>"#
            ),
        ] {
            assert!(parse_data_model(xml.as_bytes()).is_err());
        }
        assert!(parse_data_model(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
    }

    #[test]
    fn rejects_missing_connection_and_unknown_table_references() {
        let mut value = definition();
        value.tables[0].connection = "Absent".into();
        let (mut package, workbook) = fixture_package();
        assert!(
            store_data_model(
                &mut package,
                &workbook,
                &Model {
                    definition: value,
                    payload: model().payload,
                }
            )
            .is_err()
        );
        let mut value = definition();
        value.relationships[0].to_table = "Absent".into();
        assert!(write_data_model(&value).is_err());
    }

    #[test]
    fn rejects_connections_with_target_queries_or_fragments() {
        for target in ["connections.xml?query", "connections.xml#fragment"] {
            let (mut package, workbook) = fixture_package();
            replace_connections_target(&mut package, &workbook, target);
            assert!(matches!(
                store_data_model(&mut package, &workbook, &model()),
                Err(Error::Invalid(message)) if message.contains("query or fragment")
            ));
            assert!(
                package
                    .get_part(&PackURI::new(DATA_MODEL_PART_NAME).unwrap())
                    .is_err()
            );
        }
    }

    #[test]
    fn connection_name_matching_decodes_spreadsheet_escapes_once() {
        let (mut package, workbook) = fixture_package();
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        package.get_part_mut(&connections).unwrap().set_blob(
            format!(
                r#"<?xml version="1.0"?><?catalog?><connections xmlns="{SML}"><?inside?><connection id="1" name="Model_x005F_Connection" refreshedVersion="7"/></connections>"#
            )
            .into_bytes(),
        );
        let mut expected = model();
        for table in &mut expected.definition.tables {
            table.connection = "Model_Connection".into();
        }
        store_data_model(&mut package, &workbook, &expected).unwrap();
        let actual = load_data_model(&package, &workbook).unwrap().unwrap();
        assert_eq!(actual.definition.tables, expected.definition.tables);
    }

    #[test]
    fn rejects_invalid_model_connection_identity() {
        let cases = [(1, ""), (5, "non-empty"), (5, "missing-id")];
        for (connection_type, id) in cases {
            let (mut package, workbook) = fixture_package();
            let connections = PackURI::new("/xl/connections.xml").unwrap();
            let model_connection = if id == "missing-id" {
                r#"<x15:connection xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" model="true"/>"#.to_owned()
            } else {
                format!(r#"<x15:connection xmlns:x15="{X15}" model="true" id="{id}"/>"#)
            };
            package.get_part_mut(&connections).unwrap().set_blob(
                format!(
                    r#"<connections xmlns="{SML}"><connection id="1" name="ModelConnection" type="{connection_type}" refreshedVersion="7"><extLst><ext uri="{MODEL_CONNECTION_EXTENSION_URI}">{model_connection}</ext></extLst></connection></connections>"#
                )
                .into_bytes(),
            );
            assert!(matches!(
                store_data_model(&mut package, &workbook, &model()),
                Err(Error::Invalid(message)) if message.contains("model workbook connection")
            ));
        }
    }

    #[test]
    fn rejects_orphan_duplicate_wrong_path_and_relationship_edges() {
        let (mut package, workbook) = fixture_package();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(DATA_MODEL_PART_NAME).unwrap(),
            DATA_MODEL_CONTENT_TYPE.into(),
            vec![1],
        )));
        assert!(load_data_model(&package, &workbook).is_err());
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        package
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "model/item.data".into(),
                "rIdModel".into(),
                false,
            );
        assert!(load_data_model(&package, &workbook).is_err());
        let mut wrong = model();
        wrong.payload.part_name = "/xl/model/other.data".into();
        let (mut package, workbook) = fixture_package();
        assert!(store_data_model(&mut package, &workbook, &wrong).is_err());
    }

    #[test]
    fn source_bound_transaction_edits_descriptor_and_inverse_restores_exact_graph() {
        let (mut package, workbook) = fixture_package();
        let expected = model();
        store_data_model(&mut package, &workbook, &expected).unwrap();
        let before_workbook = package.get_part(&workbook).unwrap().blob().to_vec();
        let before_payload = package
            .get_part(&PackURI::new(DATA_MODEL_PART_NAME).unwrap())
            .unwrap()
            .blob()
            .to_vec();

        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(
            transaction
                .edit_definition(|definition| {
                    definition.min_version_load = 8;
                    Ok(())
                })
                .unwrap()
        );
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(
            load_data_model(&package, &workbook)
                .unwrap()
                .unwrap()
                .payload
                .data,
            before_payload
        );

        commit.patch().inverse().apply(&mut package).unwrap();
        assert_eq!(package.get_part(&workbook).unwrap().blob(), before_workbook);
        assert_eq!(
            package
                .get_part(&PackURI::new(DATA_MODEL_PART_NAME).unwrap())
                .unwrap()
                .blob(),
            before_payload
        );
    }

    #[test]
    fn source_bound_time_grouping_changed_writes_refuse_unproven_inner_identity() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut package = as_source(&package);
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Sales".into(),
                column_name: "OrderDate".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "year-1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        install_model_time_groupings(&mut package, &workbook, grouping.clone());
        let before = package.get_part(&workbook).unwrap().blob().to_vec();
        let mut changed = grouping.clone();
        changed.groupings[0].column_id = "unproven-source-column".into();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(matches!(
            transaction.edit_model_time_groupings(Some(changed)),
            Err(Error::Unsupported { feature })
                if feature.contains("inner XLDM table, column, and calculated-column identity closure")
        ));
        assert!(!transaction.is_changed());
        drop(transaction);
        assert_eq!(package.get_part(&workbook).unwrap().blob(), before);
        assert_eq!(
            load_data_model(&package, &workbook)
                .unwrap()
                .unwrap()
                .definition
                .model_time_groupings()
                .unwrap(),
            Some(grouping)
        );
    }

    #[test]
    fn source_bound_time_grouping_noop_survives_save_reopen_and_inverse() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut package = as_source(&package);
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Date".into(),
                column_name: "DateKey".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Month".into(),
                    column_id: "month-1".into(),
                    content_type: ModelTimeGroupingContentType::Months,
                    is_selected: false,
                }],
            }],
        };
        install_model_time_groupings(&mut package, &workbook, grouping.clone());
        let before = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(
            !transaction
                .edit_model_time_groupings(Some(grouping.clone()))
                .unwrap()
        );
        assert!(!transaction.is_changed());
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());

        let after = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        assert_eq!(after, before);
        let mut reopened = OpcPackage::from_bytes(&after).unwrap();
        assert_eq!(
            load_data_model(&reopened, &workbook)
                .unwrap()
                .unwrap()
                .definition
                .model_time_groupings()
                .unwrap(),
            Some(grouping)
        );
        commit.patch().inverse().apply(&mut reopened).unwrap();
        assert_eq!(
            litchi_opc::PackageWriter::to_bytes(&reopened).unwrap(),
            before
        );
    }

    #[test]
    fn source_bound_time_grouping_remove_is_reversible_without_inner_rewrite() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut package = as_source(&package);
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Sales".into(),
                column_name: "OrderDate".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "year-1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        install_model_time_groupings(&mut package, &workbook, grouping);
        let before = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(transaction.edit_model_time_groupings(None).unwrap());
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(
            load_data_model(&package, &workbook)
                .unwrap()
                .unwrap()
                .definition
                .model_time_groupings()
                .unwrap(),
            None
        );

        let after = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_bytes(&after).unwrap();
        commit.patch().inverse().apply(&mut reopened).unwrap();
        assert_eq!(
            litchi_opc::PackageWriter::to_bytes(&reopened).unwrap(),
            before
        );
    }

    #[test]
    fn source_bound_time_grouping_bad_references_never_reach_specialized_writer() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut package = as_source(&package);
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Sales".into(),
                column_name: "OrderDate".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "year-1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        install_model_time_groupings(&mut package, &workbook, grouping.clone());
        let mut candidates = Vec::new();
        let mut bad_table = grouping.clone();
        bad_table.groupings[0].table_name = "MissingTable".into();
        candidates.push(bad_table);
        let mut bad_column = grouping.clone();
        bad_column.groupings[0].column_id = "MissingColumn".into();
        candidates.push(bad_column);
        let mut bad_calculated = grouping.clone();
        bad_calculated.groupings[0].calculated_time_columns[0].column_id =
            "MissingCalculatedColumn".into();
        candidates.push(bad_calculated);
        let mut future = grouping.clone();
        future.groupings[0].calculated_time_columns[0].content_type =
            ModelTimeGroupingContentType::Other("futureUnit".into());
        candidates.push(future);

        let before = package.get_part(&workbook).unwrap().blob().to_vec();
        for candidate in candidates {
            let mut transaction = Transaction::new(&mut package).unwrap();
            assert!(matches!(
                transaction.edit_model_time_groupings(Some(candidate)),
                Err(Error::Unsupported { feature })
                    if feature.contains("inner XLDM table, column, and calculated-column identity closure")
                        || feature.contains("unknown modelTimeGrouping contentType")
            ));
            assert!(!transaction.is_changed());
        }
        assert_eq!(package.get_part(&workbook).unwrap().blob(), before);
    }

    #[test]
    fn source_bound_time_grouping_unknown_sibling_topology_is_not_rewritten() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut package = as_source(&package);
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Sales".into(),
                column_name: "OrderDate".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "year-1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        install_model_time_groupings(&mut package, &workbook, grouping.clone());
        let before = package.get_part(&workbook).unwrap().blob().to_vec();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(matches!(
            transaction.edit_definition(|definition| {
                let mut changed = grouping.clone();
                changed.groupings[0].column_id = "changed".into();
                definition.set_model_time_groupings(Some(changed))
            }),
            Err(Error::Unsupported { feature })
                if feature.contains("structural Data Model edits require validated inner XLDM identity closure")
        ));
        assert!(!transaction.is_changed());
        drop(transaction);
        assert_eq!(package.get_part(&workbook).unwrap().blob(), before);
    }

    #[test]
    fn load_rejects_unknown_model_time_grouping_table() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "MissingTable".into(),
                column_name: "OrderDate".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "year-1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        install_model_time_groupings(&mut package, &workbook, grouping);
        assert!(matches!(
            load_data_model(&package, &workbook),
            Err(Error::Invalid(message)) if message.contains("unknown table 'MissingTable'")
        ));
    }

    #[test]
    fn load_rejects_duplicate_model_time_grouping_owner() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Sales".into(),
                column_name: "OrderDate".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "year-1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        install_model_time_groupings(&mut package, &workbook, grouping);
        let source =
            String::from_utf8(package.get_part(&workbook).unwrap().blob().to_vec()).unwrap();
        let owner = format!(r#"<x15:ext uri="{MODEL_TIME_GROUPINGS_EXTENSION_URI}">"#);
        let start = source.find(&owner).unwrap();
        let end = source[start..].find("</x15:ext>").unwrap() + start + "</x15:ext>".len();
        let duplicate = source[start..end].to_owned();
        let marker = "</x15:extLst>";
        let (before_close, close) = source.rsplit_once(marker).unwrap();
        package
            .get_part_mut(&workbook)
            .unwrap()
            .set_blob(format!("{before_close}{duplicate}{marker}{close}").into_bytes());
        assert!(matches!(
            load_data_model(&package, &workbook),
            Err(Error::Invalid(message)) if message.contains("duplicate modelTimeGroupings owner")
        ));
    }

    #[test]
    fn signed_source_bound_time_grouping_write_refuses_before_commit() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Sales".into(),
                column_name: "OrderDate".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "year-1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        install_model_time_groupings(&mut package, &workbook, grouping.clone());
        let mut package = as_source(&package);
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());
        let before = package.get_part(&workbook).unwrap().blob().to_vec();
        let mut transaction = Transaction::new(&mut package).unwrap();
        let mut changed = grouping;
        changed.groupings[0].column_id = "changed".into();
        assert!(matches!(
            transaction.edit_model_time_groupings(Some(changed)),
            Err(Error::Unsupported { .. })
        ));
        assert!(!transaction.is_changed());
        drop(transaction);
        assert_eq!(package.get_part(&workbook).unwrap().blob(), before);
    }

    #[test]
    fn source_bound_transaction_removes_and_restores_the_singleton() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(transaction.remove().unwrap());
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert!(load_data_model(&package, &workbook).unwrap().is_none());
        commit.patch().inverse().apply(&mut package).unwrap();
        assert!(load_data_model(&package, &workbook).unwrap().is_some());
    }

    #[test]
    fn signed_transaction_noop_remains_allowed() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut package = as_source(&package);
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());
        let before = package.get_part(&workbook).unwrap().blob().to_vec();
        let transaction = Transaction::new(&mut package).unwrap();
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert_eq!(package.get_part(&workbook).unwrap().blob(), before);
    }

    #[test]
    fn payload_only_edit_requires_signature_provenance_disposition() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut package = as_source(&package);
        remove_low_level_signature(&mut package);

        let mut transaction = Transaction::new(&mut package).unwrap();
        let mut replacement = transaction.model().unwrap().to_owned();
        replacement.payload.data = payload_with_equivalent_object_id_case();
        assert!(transaction.set(replacement).unwrap());
        assert!(matches!(transaction.commit(), Err(Error::Signed)));

        package.unsign();
        let mut transaction = Transaction::new(&mut package).unwrap();
        let mut replacement = transaction.model().unwrap().to_owned();
        replacement.payload.data = payload_with_equivalent_object_id_case();
        assert!(transaction.set(replacement).unwrap());
        assert!(transaction.commit().is_ok());
    }

    #[test]
    fn payload_only_replacement_reproves_retained_time_groupings() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut package = as_source(&package);
        let grouping = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Sales".into(),
                column_name: "OrderDate".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "year-1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        // Install this deliberately through the low-level source fixture: it
        // lets the test model a previously published workbook whose opaque
        // XLDM bytes had already been admitted by an older reader.  The
        // ordinary edit API must reprove the retained grouping before it
        // stages a new payload.
        install_model_time_groupings(&mut package, &workbook, grouping);

        let before = package.get_part(&workbook).unwrap().blob().to_vec();
        let mut transaction = Transaction::new(&mut package).unwrap();
        let mut replacement = transaction.model().unwrap().to_owned();
        replacement.payload.data = payload_with_equivalent_object_id_case();
        assert!(matches!(
            transaction.set(replacement),
            Err(Error::Unsupported { feature })
                if feature.contains("inner XLDM table, column, and calculated-column identity closure")
        ));
        assert!(!transaction.is_changed());
        drop(transaction);
        assert_eq!(package.get_part(&workbook).unwrap().blob(), before);
    }

    #[test]
    fn snapshots_edits_publication_and_inverse_share_the_package_payload() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let uri = PackURI::new(DATA_MODEL_PART_NAME).unwrap();
        let payload = package.get_part(&uri).unwrap().blob_arc();
        let workbook_bytes = package.get_part(&workbook).unwrap().blob_arc();
        let snapshot = Snapshot::load(&package).unwrap();
        let cloned = snapshot.clone();
        assert!(Arc::ptr_eq(&snapshot.source, &cloned.source));
        assert!(Arc::ptr_eq(
            &snapshot.model().unwrap().definition,
            &cloned.model().unwrap().definition
        ));
        assert!(Arc::ptr_eq(&payload, &snapshot.model().unwrap().data));

        let mut edit = Transaction::new(&mut package).unwrap();
        assert_eq!(edit.model().unwrap().payload().as_ptr(), payload.as_ptr());
        edit.edit_definition(|definition| {
            definition.min_version_load += 1;
            Ok(())
        })
        .unwrap();
        assert_eq!(edit.model().unwrap().payload().as_ptr(), payload.as_ptr());
        let commit = edit.commit().unwrap();
        assert!(Arc::ptr_eq(
            &payload,
            &commit.snapshot().model().unwrap().data
        ));
        assert!(Arc::ptr_eq(
            &payload,
            &package.get_part(&uri).unwrap().blob_arc()
        ));
        commit.patch().inverse().apply(&mut package).unwrap();
        assert!(Arc::ptr_eq(
            &payload,
            &package.get_part(&uri).unwrap().blob_arc()
        ));
        assert!(Arc::ptr_eq(
            &workbook_bytes,
            &package.get_part(&workbook).unwrap().blob_arc()
        ));
    }

    #[test]
    fn structural_descriptor_changes_cannot_reuse_the_opaque_payload() {
        let changes: [fn(&mut Definition); 5] = [
            |definition| definition.tables[0].id = "another-id".into(),
            |definition| definition.tables[0].name = "AnotherName".into(),
            |definition| definition.tables[0].connection = "AnotherConnection".into(),
            |definition| definition.relationships.clear(),
            |definition| definition.extension_list = None,
        ];
        for change in changes {
            let (mut package, workbook) = fixture_package();
            store_data_model(&mut package, &workbook, &model()).unwrap();
            let before = package.get_part(&workbook).unwrap().blob_arc();
            let mut edit = Transaction::new(&mut package).unwrap();
            assert!(matches!(
                edit.edit_definition(|definition| {
                    change(definition);
                    Ok(())
                }),
                Err(Error::Unsupported { .. })
            ));
            assert!(!edit.is_changed());

            let mut owned = edit.model().unwrap().to_owned();
            change(&mut owned.definition);
            assert!(matches!(
                edit.set(owned.clone()),
                Err(Error::Unsupported { .. })
            ));
            assert!(!edit.is_changed());

            // Removing the draft must not hide its original payload identity.
            edit.remove().unwrap();
            assert!(matches!(edit.set(owned), Err(Error::Unsupported { .. })));
            assert!(edit.model().is_none());
            drop(edit);
            assert!(Arc::ptr_eq(
                &before,
                &package.get_part(&workbook).unwrap().blob_arc()
            ));
        }
    }

    #[test]
    fn connection_resolution_retains_case_folding_and_rejects_folded_ambiguity() {
        let (mut package, workbook) = fixture_package();
        let mut expected = model();
        for table in &mut expected.definition.tables {
            table.connection = "MODELCONNECTION".into();
        }
        store_data_model(&mut package, &workbook, &expected).unwrap();
        let actual = load_data_model(&package, &workbook).unwrap().unwrap();
        assert_eq!(actual.definition.tables, expected.definition.tables);
        assert_eq!(actual.payload, expected.payload);

        let (mut package, workbook) = fixture_package();
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        package.get_part_mut(&connections).unwrap().set_blob(format!(
            r#"<connections xmlns="{SML}"><connection id="1" name="ModelConnection"/><connection id="2" name="modelconnection"/></connections>"#,
        ).into_bytes());
        let before = Snapshot::load(&package).unwrap();
        assert!(store_data_model(&mut package, &workbook, &model()).is_err());
        assert!(Snapshot::load(&package).unwrap().same_source(&before));
    }

    #[test]
    fn duplicate_named_connections_are_ambiguous_and_publication_is_atomic() {
        let (mut package, workbook) = fixture_package();
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        package.get_part_mut(&connections).unwrap().set_blob(format!(
            r#"<connections xmlns="{SML}"><connection id="1" name="ModelConnection"/><connection id="2" name="ModelConnection"/></connections>"#,
        ).into_bytes());
        let before = package.get_part(&workbook).unwrap().blob_arc();
        assert!(store_data_model(&mut package, &workbook, &model()).is_err());
        assert!(Arc::ptr_eq(
            &before,
            &package.get_part(&workbook).unwrap().blob_arc()
        ));
        assert!(
            package
                .get_part(&PackURI::new(DATA_MODEL_PART_NAME).unwrap())
                .is_err()
        );
    }

    #[test]
    fn connections_only_changes_make_model_patches_stale() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let mut edit = Transaction::new(&mut package).unwrap();
        edit.edit_definition(|definition| {
            definition.min_version_load += 1;
            Ok(())
        })
        .unwrap();
        let commit = edit.commit().unwrap();
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        let content = package.get_part(&connections).unwrap().blob().to_vec();
        let mut changed = String::from_utf8(content).unwrap();
        changed = changed.replace("refreshedVersion=\"7\"", "refreshedVersion=\"8\"");
        package
            .get_part_mut(&connections)
            .unwrap()
            .set_blob(changed.into_bytes());
        let before = package.get_part(&workbook).unwrap().blob_arc();
        assert!(matches!(
            commit.patch().inverse().apply(&mut package),
            Err(Error::PatchConflict { .. })
        ));
        assert!(Arc::ptr_eq(
            &before,
            &package.get_part(&workbook).unwrap().blob_arc()
        ));
    }

    #[test]
    fn load_version_edit_splices_only_the_attribute_value() {
        let (mut package, workbook) = fixture_package();
        store_data_model(&mut package, &workbook, &model()).unwrap();
        let original =
            String::from_utf8(package.get_part(&workbook).unwrap().blob().to_vec()).unwrap();
        let original = original
            .replace("x15:", "d:")
            .replace("xmlns:x15", "xmlns:d")
            .replace("minVersionLoad=\"7\"", "minVersionLoad = '07'")
            .replace("<d:modelTables>", "<!--before tables--><d:modelTables>")
            .replace("</d:modelTable>", "</d:modelTable><!--after table-->")
            .replace(
                "<v:opaque xmlns:v=\"urn:vendor\"/>",
                "<v:opaque xmlns:v=\"urn:vendor\" q=\"v:Value\"><!--inside extension--></v:opaque>",
            );
        package
            .get_part_mut(&workbook)
            .unwrap()
            .set_blob(original.as_bytes().to_vec());
        let mut package = as_source(&package);
        let mut edit = Transaction::new(&mut package).unwrap();
        edit.edit_definition(|definition| {
            definition.min_version_load = 8;
            Ok(())
        })
        .unwrap();
        let commit = edit.commit().unwrap();
        let saved = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let reopened = OpcPackage::from_bytes(&saved).unwrap();
        assert!(
            Snapshot::load(&reopened)
                .unwrap()
                .same_source(commit.snapshot())
        );

        assert_eq!(
            package.get_part(&workbook).unwrap().blob(),
            original
                .replace("minVersionLoad = '07'", "minVersionLoad = '8'")
                .as_bytes()
        );
        commit.patch().inverse().apply(&mut package).unwrap();
        let restored =
            OpcPackage::from_bytes(&litchi_opc::PackageWriter::to_bytes(&package).unwrap())
                .unwrap();
        assert_eq!(
            restored.get_part(&workbook).unwrap().blob(),
            original.as_bytes()
        );

        assert_eq!(
            package.get_part(&workbook).unwrap().blob(),
            original.as_bytes()
        );
    }

    #[test]
    fn adding_a_load_version_to_an_empty_descriptor_preserves_its_spelling() {
        for descriptor in ["<d:dataModel />", "<d:dataModel><!--keep--></d:dataModel>"] {
            let (mut package, workbook) = fixture_package();
            let mut value = model();
            value.definition = Definition::default();
            store_data_model(&mut package, &workbook, &value).unwrap();
            let original = format!(
                r#"<workbook xmlns="{SML}" xmlns:d="{X15}"><sheets/><extLst><ext uri="{DATA_MODEL_EXTENSION_URI}">{descriptor}</ext></extLst></workbook>"#
            );
            package
                .get_part_mut(&workbook)
                .unwrap()
                .set_blob(original.as_bytes().to_vec());
            let mut package = as_source(&package);
            let mut edit = Transaction::new(&mut package).unwrap();
            edit.edit_definition(|definition| {
                definition.min_version_load = 6;
                Ok(())
            })
            .unwrap();
            let commit = edit.commit().unwrap();
            let saved = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
            let reopened = OpcPackage::from_bytes(&saved).unwrap();
            assert!(
                Snapshot::load(&reopened)
                    .unwrap()
                    .same_source(commit.snapshot())
            );

            assert_eq!(
                commit
                    .snapshot()
                    .model()
                    .unwrap()
                    .definition()
                    .min_version_load,
                6
            );
            commit.patch().inverse().apply(&mut package).unwrap();
            let restored =
                OpcPackage::from_bytes(&litchi_opc::PackageWriter::to_bytes(&package).unwrap())
                    .unwrap();
            assert_eq!(
                restored.get_part(&workbook).unwrap().blob(),
                original.as_bytes()
            );

            assert_eq!(
                package.get_part(&workbook).unwrap().blob(),
                original.as_bytes()
            );
        }
    }
    #[test]
    fn descriptor_graph_creation_removal_and_inverse_preserve_noncompact_workbooks() {
        for list in [
            "",
            "<s:extLst />",
            "<s:extLst><!--keep--></s:extLst>",
            "<s:extLst><s:ext uri='urn:foreign'><v:keep xmlns:v='urn:vendor' /></s:ext><!--keep--></s:extLst>",
        ] {
            let (mut package, workbook) = fixture_package();
            let original =
                format!("<s:workbook xmlns:s='{SML}' ><!--root--><s:sheets/>{list}</s:workbook>");
            package
                .get_part_mut(&workbook)
                .unwrap()
                .set_blob(original.as_bytes().to_vec());
            let mut package = as_source(&package);
            let before = Snapshot::load(&package).unwrap();
            let mut value = model();
            value.definition.extension_list = None;
            let mut edit = Transaction::new(&mut package).unwrap();
            edit.set(value).unwrap();
            let inserted = edit.commit().unwrap();
            let saved = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
            let reopened = OpcPackage::from_bytes(&saved).unwrap();
            assert!(
                Snapshot::load(&reopened)
                    .unwrap()
                    .same_source(inserted.snapshot())
            );
            assert!(
                String::from_utf8_lossy(reopened.get_part(&workbook).unwrap().blob())
                    .contains("<!--root-->")
            );
            inserted.patch().inverse().apply(&mut package).unwrap();
            let restored =
                OpcPackage::from_bytes(&litchi_opc::PackageWriter::to_bytes(&package).unwrap())
                    .unwrap();
            assert!(Snapshot::load(&restored).unwrap().same_source(&before));
            inserted.patch().apply(&mut package).unwrap();
            let mut remove = Transaction::new(&mut package).unwrap();
            remove.remove().unwrap();
            let removed = remove.commit().unwrap();
            let reopened =
                OpcPackage::from_bytes(&litchi_opc::PackageWriter::to_bytes(&package).unwrap())
                    .unwrap();
            assert!(Snapshot::load(&reopened).unwrap().is_empty());
            removed.patch().inverse().apply(&mut package).unwrap();
            let restored =
                OpcPackage::from_bytes(&litchi_opc::PackageWriter::to_bytes(&package).unwrap())
                    .unwrap();
            assert!(
                Snapshot::load(&restored)
                    .unwrap()
                    .same_source(inserted.snapshot())
            );
        }
    }

    #[test]
    fn load_version_accepts_xml_schema_integer_whitespace_but_not_other_spacing() {
        for lexical in ["  +007  ", "&#x9;005&#xA;"] {
            let xml = format!(r#"<d:dataModel xmlns:d="{X15}" minVersionLoad="{lexical}"/>"#);
            assert!(matches!(
                parse_data_model(xml.as_bytes()).unwrap().min_version_load,
                5 | 7
            ));
        }
        for lexical in ["4", "256", "5.0", "&#xA0;5", "5 0"] {
            let xml = format!(r#"<d:dataModel xmlns:d="{X15}" minVersionLoad="{lexical}"/>"#);
            assert!(parse_data_model(xml.as_bytes()).is_err());
        }
    }
}
