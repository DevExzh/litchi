//! Bounded package graphs owned by standard DrawingML Chart parts.
//!
//! XLSB uses the same Chart, chart user-shapes, embedded-package, OLE, and
//! relationship grammars as XLSX. This module validates, authors, and resolves
//! those inert resources without opening linked targets or embedded payloads.

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsx::{
    ChartExternalDataPart, ChartExternalDataTarget, ChartRelationship, ChartRelationshipTarget,
    ChartUserShapesPart, WorksheetChart,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rel};
use litchi_opc::part::Part;
use litchi_opc::{BlobPart, ContentType, OpcPackage, PackURI};
use std::collections::HashSet;

const MAX_RELATIONSHIPS: usize = 4_096;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1_024;
const MAX_RELATIONSHIP_TYPE_BYTES: usize = 4_096;
const MAX_EXTERNAL_TARGET_BYTES: usize = 32_767;
const MAX_EXTENSION_BYTES: usize = 16;
const MAX_RESOURCE_BYTES: usize = 64 * 1024 * 1024;
const MAX_CHART_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_USER_SHAPES_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_TOTAL_RESOURCE_BYTES: usize = 256 * 1024 * 1024;

pub(crate) struct AuthoredChartGraph {
    pub chart_part: BlobPart,
    pub related_parts: Vec<BlobPart>,
}

pub(crate) struct ResolvedChartGraph {
    pub chart: crate::charts::Chart,
    pub external_data_part: Option<ChartExternalDataPart>,
    pub user_shapes_part: Option<ChartUserShapesPart>,
    pub additional_relationships: Vec<ChartRelationship>,
}

/// Validate a chart's complete package-resource model without mutating it.
pub(crate) fn validate_chart_resources(chart: &WorksheetChart) -> XlsbResult<()> {
    let (external_id, user_shapes_id) = validate_chart_resource_metadata(chart)?;
    validated_chart_xml(chart, external_id.as_deref(), user_shapes_id.as_deref())?;
    Ok(())
}

fn validate_chart_resource_metadata(
    chart: &WorksheetChart,
) -> XlsbResult<(Option<String>, Option<String>)> {
    if chart.chart.external_data.is_some() != chart.external_data_part.is_some() {
        return Err(invalid(
            "chart external-data metadata and package payload disagree",
        ));
    }
    if chart.chart.user_shapes.is_some() != chart.user_shapes_part.is_some() {
        return Err(invalid(
            "chart user-shapes metadata and package payload disagree",
        ));
    }
    let relationship_count = chart
        .additional_relationships
        .len()
        .checked_add(
            chart
                .user_shapes_part
                .as_ref()
                .map_or(0, |value| value.relationships.len()),
        )
        .and_then(|value| value.checked_add(usize::from(chart.external_data_part.is_some())))
        .and_then(|value| value.checked_add(usize::from(chart.user_shapes_part.is_some())))
        .ok_or_else(|| limit("chart relationship count"))?;
    if relationship_count > MAX_RELATIONSHIPS {
        return Err(limit("chart relationship count"));
    }

    let mut total = 0usize;
    let mut chart_ids = HashSet::with_capacity(
        chart.additional_relationships.len()
            + usize::from(chart.external_data_part.is_some())
            + usize::from(chart.user_shapes_part.is_some()),
    );
    for relationship in &chart.additional_relationships {
        validate_relationship(relationship, &mut total)?;
        if !chart_ids.insert(relationship.relationship_id.clone()) {
            return Err(invalid(format!(
                "duplicate chart relationship ID {:?}",
                relationship.relationship_id
            )));
        }
    }

    if let Some(external) = &chart.external_data_part {
        validate_external_data(external, &mut total)?;
        let id = chart
            .chart
            .external_data
            .as_ref()
            .and_then(|value| value.relationship_id.as_deref())
            .map(str::to_string)
            .unwrap_or_else(|| next_relationship_id(&chart_ids));
        validate_relationship_id(&id)?;
        if !chart_ids.insert(id.clone()) {
            return Err(invalid(format!(
                "conflicting chart external-data relationship ID {id:?}"
            )));
        }
    }

    if let Some(user_shapes) = &chart.user_shapes_part {
        validate_user_shapes(user_shapes, &mut total)?;
        let id = chart
            .chart
            .user_shapes
            .as_ref()
            .and_then(|value| value.relationship_id.as_deref())
            .map(str::to_string)
            .unwrap_or_else(|| next_relationship_id(&chart_ids));
        validate_relationship_id(&id)?;
        if !chart_ids.insert(id.clone()) {
            return Err(invalid(format!(
                "conflicting chart user-shapes relationship ID {id:?}"
            )));
        }
    }

    for id in crate::xlsx::chart::chart_fragment_relationship_ids(&chart.chart)? {
        if !chart_ids.contains(id.as_str()) {
            return Err(invalid(format!(
                "chart fragment references undeclared relationship {id:?}"
            )));
        }
    }
    planned_special_relationship_ids(chart)
}

/// Build one Chart part and all of its internal related resource parts.
pub(crate) fn author_chart_graph(
    chart: &WorksheetChart,
    chart_index: usize,
) -> XlsbResult<AuthoredChartGraph> {
    let (planned_external_id, planned_user_shapes_id) = validate_chart_resource_metadata(chart)?;
    let chart_name = format!("chart{chart_index}.xml");
    let mut chart_part = BlobPart::new(
        PackURI::new(format!("/xl/charts/{chart_name}"))?,
        ct::DML_CHART.to_string(),
        Vec::new(),
    );
    let mut related_parts = Vec::new();

    for (ordinal, relationship) in chart.additional_relationships.iter().enumerate() {
        let (target, external) = author_relationship_target(
            &relationship.target,
            "chartResources",
            &format!("chartResource{chart_index}_{}", ordinal + 1),
            &mut related_parts,
        )?;
        chart_part.rels_mut().add_relationship(
            relationship.relationship_type.clone(),
            target,
            relationship.relationship_id.clone(),
            external,
        );
    }

    let external_data_id = if let Some(external) = &chart.external_data_part {
        let (target, is_external) = match &external.target {
            ChartExternalDataTarget::Embedded {
                data,
                content_type,
                extension,
            } => {
                let filename = format!("chartData{chart_index}.{}", extension.to_ascii_lowercase());
                related_parts.push(BlobPart::new(
                    PackURI::new(format!("/xl/embeddings/{filename}"))?,
                    content_type.clone(),
                    data.clone(),
                ));
                (format!("../embeddings/{filename}"), false)
            },
            ChartExternalDataTarget::Linked { target } => (target.clone(), true),
        };
        Some(add_chart_relationship(
            &mut chart_part,
            planned_external_id
                .as_deref()
                .ok_or_else(|| invalid("chart external-data relationship ID was not planned"))?,
            &external.relationship_type,
            &target,
            is_external,
        ))
    } else {
        None
    };

    let user_shapes_id = if let Some(user_shapes) = &chart.user_shapes_part {
        let filename = format!("chartDrawing{chart_index}.xml");
        let mut part = BlobPart::new(
            PackURI::new(format!("/xl/drawings/{filename}"))?,
            ct::DML_CHARTSHAPES.to_string(),
            user_shapes.xml.clone(),
        );
        for (ordinal, relationship) in user_shapes.relationships.iter().enumerate() {
            let (target, external) = author_relationship_target(
                &relationship.target,
                "media",
                &format!("chartShape{chart_index}_{}", ordinal + 1),
                &mut related_parts,
            )?;
            part.rels_mut().add_relationship(
                relationship.relationship_type.clone(),
                target,
                relationship.relationship_id.clone(),
                external,
            );
        }
        related_parts.push(part);
        Some(add_chart_relationship(
            &mut chart_part,
            planned_user_shapes_id
                .as_deref()
                .ok_or_else(|| invalid("chart user-shapes relationship ID was not planned"))?,
            rel::CHART_USER_SHAPES,
            &format!("../drawings/{filename}"),
            false,
        ))
    } else {
        None
    };

    let xml = validated_chart_xml(
        chart,
        external_data_id.as_deref(),
        user_shapes_id.as_deref(),
    )?;
    chart_part.set_blob(xml);
    Ok(AuthoredChartGraph {
        chart_part,
        related_parts,
    })
}

/// Parse one Chart part and resolve its complete resource graph.
pub(crate) fn parse_chart_resources(
    package: &OpcPackage,
    chart_part: &dyn Part,
) -> XlsbResult<ResolvedChartGraph> {
    if chart_part.blob().len() > MAX_CHART_XML_BYTES {
        return Err(limit("chart XML bytes"));
    }
    let chart_relationship_count = chart_part.rels().iter().count();
    if chart_relationship_count > MAX_RELATIONSHIPS {
        return Err(limit("chart relationship count"));
    }
    let chart = crate::charts::reader::parse_chart(chart_part.blob())?;
    let mut total = 0usize;
    let mut consumed = HashSet::new();

    let external_data = if let Some(metadata) = &chart.external_data {
        let id = metadata
            .relationship_id
            .as_deref()
            .ok_or_else(|| invalid("parsed chart external data has no relationship ID"))?;
        validate_relationship_id(id)?;
        let relationship = chart_part.rels().get(id).ok_or_else(|| {
            invalid(format!(
                "chart external data references missing relationship {id:?}"
            ))
        })?;
        if !crate::xlsx::chart::is_chart_external_data_relationship_type(relationship.reltype()) {
            return Err(invalid(format!(
                "chart external-data relationship {id:?} has invalid type {:?}",
                relationship.reltype()
            )));
        }
        consumed.insert(id.to_string());
        let target = if relationship.is_external() {
            validate_external_target(relationship.target_ref())?;
            ChartExternalDataTarget::Linked {
                target: relationship.target_ref().to_string(),
            }
        } else {
            let part = package.get_part(&relationship.target_partname()?)?;
            let expected =
                crate::xlsx::chart::chart_external_data_content_type(relationship.reltype())
                    .expect("external-data relationship type was validated");
            if part.content_type() != expected || part.rels().iter().next().is_some() {
                return Err(invalid(
                    "embedded chart external-data part has invalid content type or relationships",
                ));
            }
            add_resource_bytes(&mut total, part.blob().len())?;
            ChartExternalDataTarget::Embedded {
                data: part.blob().to_vec(),
                content_type: part.content_type().to_string(),
                extension: validated_part_extension(part.partname())?,
            }
        };
        Some(ChartExternalDataPart {
            relationship_type: relationship.reltype().to_string(),
            target,
        })
    } else {
        None
    };

    let user_shapes = if let Some(metadata) = &chart.user_shapes {
        let id = metadata
            .relationship_id
            .as_deref()
            .ok_or_else(|| invalid("parsed chart user shapes have no relationship ID"))?;
        validate_relationship_id(id)?;
        let relationship = chart_part.rels().get(id).ok_or_else(|| {
            invalid(format!(
                "chart user shapes reference missing relationship {id:?}"
            ))
        })?;
        if !crate::xlsx::chart::is_chart_user_shapes_relationship_type(relationship.reltype())
            || relationship.is_external()
        {
            return Err(invalid(format!(
                "chart user-shapes relationship {id:?} has invalid type or target mode"
            )));
        }
        consumed.insert(id.to_string());
        let part = package.get_part(&relationship.target_partname()?)?;
        if part.content_type() != ct::DML_CHARTSHAPES
            || part.blob().len() > MAX_USER_SHAPES_XML_BYTES
            || part.rels().iter().count() > MAX_RELATIONSHIPS
        {
            return Err(invalid("invalid or oversized chart user-shapes part"));
        }
        if chart_relationship_count
            .checked_add(part.rels().iter().count())
            .is_none_or(|count| count > MAX_RELATIONSHIPS)
        {
            return Err(limit("combined chart relationship count"));
        }
        add_resource_bytes(&mut total, part.blob().len())?;
        let referenced = crate::xlsx::chart::chart_user_shapes_relationship_ids(part.blob())?;
        if referenced.len() != part.rels().iter().count()
            || !referenced.iter().all(|id| part.rels().get(id).is_some())
        {
            return Err(invalid("chart user-shapes XML and relationships disagree"));
        }
        let mut ids = referenced.into_iter().collect::<Vec<_>>();
        ids.sort_unstable();
        let mut relationships = Vec::with_capacity(ids.len());
        for id in ids {
            let relationship = part
                .rels()
                .get(&id)
                .expect("user-shapes relationship presence was validated");
            relationships.push(resolve_relationship(package, relationship, &mut total)?);
        }
        Some(ChartUserShapesPart {
            xml: part.blob().to_vec(),
            relationships,
        })
    } else {
        None
    };

    let mut additional = Vec::new();
    let mut remaining = chart_part
        .rels()
        .iter()
        .filter(|relationship| !consumed.contains(relationship.r_id()))
        .collect::<Vec<_>>();
    remaining.sort_unstable_by_key(|relationship| relationship.r_id());
    for relationship in remaining {
        additional.push(resolve_relationship(package, relationship, &mut total)?);
    }
    let fragment_ids = crate::xlsx::chart::chart_fragment_relationship_ids(&chart)?;
    if !fragment_ids.iter().all(|id| {
        additional
            .iter()
            .any(|relationship| relationship.relationship_id == *id)
            || consumed.contains(id)
    }) {
        return Err(invalid(
            "chart fragment references a missing package relationship",
        ));
    }
    Ok(ResolvedChartGraph {
        chart,
        external_data_part: external_data,
        user_shapes_part: user_shapes,
        additional_relationships: additional,
    })
}

fn validate_external_data(value: &ChartExternalDataPart, total: &mut usize) -> XlsbResult<()> {
    validate_relationship_type(&value.relationship_type)?;
    let expected = crate::xlsx::chart::chart_external_data_content_type(&value.relationship_type)
        .ok_or_else(|| {
        invalid(format!(
            "unsupported chart external-data relationship type {:?}",
            value.relationship_type
        ))
    })?;
    match &value.target {
        ChartExternalDataTarget::Embedded {
            data,
            content_type,
            extension,
        } => {
            validate_embedded(data, content_type, extension, total)?;
            if content_type != expected {
                return Err(invalid(format!(
                    "chart external-data content type {content_type:?} does not match its relationship type"
                )));
            }
        },
        ChartExternalDataTarget::Linked { target } => validate_external_target(target)?,
    }
    Ok(())
}

fn validate_user_shapes(value: &ChartUserShapesPart, total: &mut usize) -> XlsbResult<()> {
    if value.xml.len() > MAX_USER_SHAPES_XML_BYTES {
        return Err(limit("chart user-shapes XML bytes"));
    }
    add_resource_bytes(total, value.xml.len())?;
    let referenced = crate::xlsx::chart::chart_user_shapes_relationship_ids(&value.xml)?;
    let mut declared = HashSet::with_capacity(value.relationships.len());
    for relationship in &value.relationships {
        validate_relationship(relationship, total)?;
        if !declared.insert(relationship.relationship_id.as_str()) {
            return Err(invalid(format!(
                "duplicate chart user-shapes relationship ID {:?}",
                relationship.relationship_id
            )));
        }
    }
    if referenced.len() != declared.len()
        || !referenced.iter().all(|id| declared.contains(id.as_str()))
    {
        return Err(invalid(
            "chart user-shapes XML and relationship declarations disagree",
        ));
    }
    Ok(())
}

fn validate_relationship(value: &ChartRelationship, total: &mut usize) -> XlsbResult<()> {
    validate_relationship_id(&value.relationship_id)?;
    validate_relationship_type(&value.relationship_type)?;
    match &value.target {
        ChartRelationshipTarget::Embedded {
            data,
            content_type,
            extension,
        } => validate_embedded(data, content_type, extension, total),
        ChartRelationshipTarget::External { target } => validate_external_target(target),
    }
}

fn validate_embedded(
    data: &[u8],
    content_type: &str,
    extension: &str,
    total: &mut usize,
) -> XlsbResult<()> {
    if data.len() > MAX_RESOURCE_BYTES {
        return Err(limit("individual chart resource bytes"));
    }
    ContentType::new(content_type.to_string())?;
    validate_extension(extension)?;
    add_resource_bytes(total, data.len())
}

fn validate_relationship_id(value: &str) -> XlsbResult<()> {
    if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid("invalid chart relationship ID length"));
    }
    let mut bytes = value.bytes();
    let first = bytes.next().expect("empty ID was rejected");
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return Err(invalid(format!(
            "chart relationship ID {value:?} is not an XML NCName"
        )));
    }
    Ok(())
}

fn validate_relationship_type(value: &str) -> XlsbResult<()> {
    if value.is_empty()
        || value.len() > MAX_RELATIONSHIP_TYPE_BYTES
        || !value.contains(':')
        || value
            .chars()
            .any(|character| character.is_control() || character.is_whitespace())
    {
        return Err(invalid(format!(
            "invalid chart relationship type URI {value:?}"
        )));
    }
    Ok(())
}

fn validate_external_target(value: &str) -> XlsbResult<()> {
    if value.is_empty()
        || value.len() > MAX_EXTERNAL_TARGET_BYTES
        || value
            .chars()
            .any(|character| character.is_control() || character.is_whitespace())
    {
        return Err(invalid("invalid chart external relationship target"));
    }
    Ok(())
}

fn validate_extension(value: &str) -> XlsbResult<()> {
    if value.is_empty()
        || value.len() > MAX_EXTENSION_BYTES
        || !value.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(invalid(format!(
            "invalid chart resource extension {value:?}"
        )));
    }
    Ok(())
}

fn add_resource_bytes(total: &mut usize, size: usize) -> XlsbResult<()> {
    if size > MAX_RESOURCE_BYTES {
        return Err(limit("individual chart resource bytes"));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("total chart resource bytes"))?;
    if *total > MAX_TOTAL_RESOURCE_BYTES {
        return Err(limit("total chart resource bytes"));
    }
    Ok(())
}

fn author_relationship_target(
    target: &ChartRelationshipTarget,
    directory: &str,
    stem: &str,
    parts: &mut Vec<BlobPart>,
) -> XlsbResult<(String, bool)> {
    match target {
        ChartRelationshipTarget::Embedded {
            data,
            content_type,
            extension,
        } => {
            let filename = format!("{stem}.{}", extension.to_ascii_lowercase());
            parts.push(BlobPart::new(
                PackURI::new(format!("/xl/{directory}/{filename}"))?,
                content_type.clone(),
                data.clone(),
            ));
            Ok((format!("../{directory}/{filename}"), false))
        },
        ChartRelationshipTarget::External { target } => Ok((target.clone(), true)),
    }
}

fn planned_special_relationship_ids(
    chart: &WorksheetChart,
) -> XlsbResult<(Option<String>, Option<String>)> {
    let mut used = chart
        .additional_relationships
        .iter()
        .map(|relationship| relationship.relationship_id.clone())
        .collect::<HashSet<_>>();
    let external = if chart.external_data_part.is_some() {
        let id = chart
            .chart
            .external_data
            .as_ref()
            .and_then(|value| value.relationship_id.clone())
            .unwrap_or_else(|| next_relationship_id(&used));
        if !used.insert(id.clone()) {
            return Err(invalid(format!(
                "conflicting chart external-data relationship ID {id:?}"
            )));
        }
        Some(id)
    } else {
        None
    };
    let user_shapes = if chart.user_shapes_part.is_some() {
        let id = chart
            .chart
            .user_shapes
            .as_ref()
            .and_then(|value| value.relationship_id.clone())
            .unwrap_or_else(|| next_relationship_id(&used));
        if !used.insert(id.clone()) {
            return Err(invalid(format!(
                "conflicting chart user-shapes relationship ID {id:?}"
            )));
        }
        Some(id)
    } else {
        None
    };
    Ok((external, user_shapes))
}

fn validated_chart_xml(
    chart: &WorksheetChart,
    external_data_id: Option<&str>,
    user_shapes_id: Option<&str>,
) -> XlsbResult<Vec<u8>> {
    let xml = crate::xlsx::chart::generate_chart_xml_with_external_data_id(
        &chart.chart,
        external_data_id,
        user_shapes_id,
    )?;
    if xml.len() > MAX_CHART_XML_BYTES {
        return Err(limit("chart XML bytes"));
    }
    crate::charts::reader::parse_chart(xml.as_slice())?;
    Ok(xml)
}

fn next_relationship_id(used: &HashSet<String>) -> String {
    let mut index = 1usize;
    loop {
        let candidate = format!("rId{index}");
        if !used.contains(&candidate) {
            return candidate;
        }
        index += 1;
    }
}

fn add_chart_relationship(
    chart_part: &mut BlobPart,
    relationship_id: &str,
    relationship_type: &str,
    target: &str,
    external: bool,
) -> String {
    chart_part.rels_mut().add_relationship(
        relationship_type.to_string(),
        target.to_string(),
        relationship_id.to_string(),
        external,
    );
    relationship_id.to_string()
}

fn resolve_relationship(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    total: &mut usize,
) -> XlsbResult<ChartRelationship> {
    validate_relationship_id(relationship.r_id())?;
    validate_relationship_type(relationship.reltype())?;
    let target = if relationship.is_external() {
        validate_external_target(relationship.target_ref())?;
        ChartRelationshipTarget::External {
            target: relationship.target_ref().to_string(),
        }
    } else {
        let part = package.get_part(&relationship.target_partname()?)?;
        if part.rels().iter().next().is_some() {
            return Err(invalid(format!(
                "chart resource part {:?} has nested relationships",
                part.partname()
            )));
        }
        add_resource_bytes(total, part.blob().len())?;
        ChartRelationshipTarget::Embedded {
            data: part.blob().to_vec(),
            content_type: part.content_type().to_string(),
            extension: validated_part_extension(part.partname())?,
        }
    };
    Ok(ChartRelationship {
        relationship_id: relationship.r_id().to_string(),
        relationship_type: relationship.reltype().to_string(),
        target,
    })
}

fn validated_part_extension(uri: &PackURI) -> XlsbResult<String> {
    let extension = uri.ext();
    validate_extension(extension)?;
    Ok(extension.to_string())
}

fn invalid(message: impl Into<String>) -> XlsbError {
    XlsbError::InvalidFormula(message.into())
}

fn limit(what: &str) -> XlsbError {
    XlsbError::InvalidFormula(format!(
        "{what} exceeds the XLSB chart-resource safety limit"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::ChartExternalData;
    use crate::xlsx::{ChartAnchor, WorksheetChart};

    #[test]
    fn reader_refuses_missing_external_data_relationship() {
        let mut worksheet_chart = WorksheetChart::bar_chart(
            "Missing",
            "Data!$A$1:$A$2",
            "Data!$B$1:$B$2",
            ChartAnchor::new(0, 0, 5, 5),
        )
        .unwrap();
        worksheet_chart.chart.external_data = Some(ChartExternalData::new("rId1"));
        let xml = crate::xlsx::chart::generate_chart_xml(&worksheet_chart.chart).unwrap();
        let uri = PackURI::new("/xl/charts/chart1.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            ct::DML_CHART.to_string(),
            xml,
        )));
        let part = package.get_part(&uri).unwrap();
        assert!(parse_chart_resources(&package, part).is_err());
    }
}
