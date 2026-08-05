//! OPC relationship graph loading and deterministic snapshot storage.

use super::codec::{
    add_total, document_references, invalid, is_chart_rel, limit, ownership, relationship_target,
    require_content_type, scan_chart_xml, validate_companion_xml, validate_graph_value,
    validate_id, validate_leaf_path,
};
use super::model::{
    CHART_CT, COLOR_STYLE_CT, COLOR_STYLE_REL, Companion, DOCUMENT_CT,
    EmbeddedWorkbook, EmbeddedWorkbookContentType, Graph, MAX_CHARTS, MAX_COMPANION_XML,
    MAX_COMPANIONS, MAX_RELATIONSHIPS, MAX_WORKBOOK_BYTES, Resource, STYLE_CT, STYLE_REL,
};
use crate::error::{Error, Result};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};
use std::collections::{BTreeSet, HashSet};

/// Load the complete bounded chart graph owned by a DOCX main document.
pub fn load(package: &OpcPackage, document_name: &PackURI) -> Result<Graph> {
    let document = package.get_part(document_name)?;
    if document.content_type() != DOCUMENT_CT {
        return Err(invalid(
            "classic chart graph requires a macro-free DOCX main part",
        ));
    }
    let (conformance, references) = document_references(document.blob())?;
    if references.len() > MAX_CHARTS {
        return Err(limit("chart count"));
    }
    let reference_set: BTreeSet<_> = references.iter().cloned().collect();
    if reference_set.len() != references.len() {
        return Err(invalid(
            "document chart relationship references are duplicated",
        ));
    }
    let chart_relationships: Vec<_> = document
        .rels()
        .iter()
        .filter(|relationship| is_chart_rel(relationship.reltype()))
        .collect();
    if chart_relationships.len() != references.len() {
        return Err(invalid(
            "document chart references and relationships differ",
        ));
    }
    let mut charts = Vec::with_capacity(references.len());
    let mut total = 0usize;
    let mut discovered_charts = BTreeSet::new();
    let mut discovered_styles = BTreeSet::new();
    let mut discovered_colors = BTreeSet::new();
    for reference in references {
        validate_id(&reference)?;
        let relationship = document
            .rels()
            .get(&reference)
            .ok_or_else(|| invalid("document chart relationship is missing"))?;
        if relationship.reltype() != conformance.chart_rel() || relationship.is_external() {
            return Err(invalid(
                "document chart relationship has wrong type or target mode",
            ));
        }
        let chart_name = relationship_target(document, relationship)?;
        validate_leaf_path(&chart_name, "/word/charts/", "chart")?;
        if !discovered_charts.insert(chart_name.as_str().to_owned()) {
            return Err(invalid(
                "multiple document anchors reference the same chart",
            ));
        }
        let chart_part = package.get_part(&chart_name)?;
        require_content_type(chart_part, CHART_CT, "chart")?;
        let scan = scan_chart_xml(chart_part.blob(), conformance)?;
        add_total(
            &mut total,
            chart_part.blob().len(),
            super::model::MAX_CHART_XML,
            "chart bytes",
        )?;
        if chart_part.rels().iter().count() > MAX_RELATIONSHIPS {
            return Err(limit("chart relationship count"));
        }
        let mut styles = Vec::new();
        let mut color_styles = Vec::new();
        let mut workbook = None;
        let mut ids = HashSet::new();
        for child in chart_part.rels().iter() {
            validate_id(child.r_id())?;
            if !ids.insert(child.r_id()) {
                return Err(invalid("chart relationship IDs collide"));
            }
            if child.is_external() {
                return Err(invalid("external chart relationship is rejected"));
            }
            match child.reltype() {
                STYLE_REL => {
                    if styles.len() >= MAX_COMPANIONS {
                        return Err(limit("chart-style count"));
                    }
                    let resource = load_companion(
                        package,
                        chart_part,
                        child,
                        STYLE_CT,
                        "chartStyle",
                        "chart style",
                        &mut total,
                    )?;
                    if !discovered_styles.insert(resource.part_name.clone()) {
                        return Err(invalid("chart-style part is shared or duplicated"));
                    }
                    styles.push(resource);
                },
                COLOR_STYLE_REL => {
                    if color_styles.len() >= MAX_COMPANIONS {
                        return Err(limit("chart-color-style count"));
                    }
                    let resource = load_companion(
                        package,
                        chart_part,
                        child,
                        COLOR_STYLE_CT,
                        "colorStyle",
                        "chart color style",
                        &mut total,
                    )?;
                    if !discovered_colors.insert(resource.part_name.clone()) {
                        return Err(invalid("chart-color-style part is shared or duplicated"));
                    }
                    color_styles.push(resource);
                },
                value if value == conformance.package_rel() => {
                    if workbook.is_some() {
                        return Err(invalid(
                            "chart has multiple embedded workbook relationships",
                        ));
                    }
                    let target = relationship_target(chart_part, child)?;
                    validate_leaf_path(&target, "/word/embeddings/", "embedded workbook")?;
                    let part = package.get_part(&target)?;
                    let content_type = EmbeddedWorkbookContentType::parse(part.content_type())
                        .ok_or_else(|| {
                            invalid(
                                "embedded chart workbook has invalid or macro-enabled content type",
                            )
                        })?;
                    if !content_type.validates_path(target.as_str()) {
                        return Err(invalid(
                            "embedded chart workbook content type and suffix differ",
                        ));
                    }
                    if part.rels().iter().next().is_some() {
                        return Err(invalid("embedded chart workbook is not an opaque leaf"));
                    }
                    add_total(
                        &mut total,
                        part.blob().len(),
                        MAX_WORKBOOK_BYTES,
                        "embedded workbook bytes",
                    )?;
                    workbook = Some(EmbeddedWorkbook {
                        relationship_id: child.r_id().to_owned(),
                        part_name: target.as_str().to_owned(),
                        content_type,
                        data: part.blob().to_vec(),
                    });
                },
                _ => return Err(invalid("chart has an unsupported nested relationship")),
            }
        }
        let actual_workbook = workbook
            .as_ref()
            .map(|value| value.relationship_id.as_str());
        if scan.workbook_id.as_deref() != actual_workbook {
            return Err(invalid(
                "chart externalData and embedded workbook relationship differ",
            ));
        }
        styles.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
        color_styles.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
        charts.push(Resource {
            document_relationship_id: reference,
            part_name: chart_name.as_str().to_owned(),
            content_type: chart_part.content_type().to_owned(),
            data: chart_part.blob().to_vec(),
            styles,
            color_styles,
            workbook,
        });
    }
    if package
        .iter_parts()
        .filter(|part| part.content_type() == CHART_CT)
        .any(|part| !discovered_charts.contains(part.partname().as_str()))
        || package
            .iter_parts()
            .filter(|part| part.content_type() == STYLE_CT)
            .any(|part| !discovered_styles.contains(part.partname().as_str()))
        || package
            .iter_parts()
            .filter(|part| part.content_type() == COLOR_STYLE_CT)
            .any(|part| !discovered_colors.contains(part.partname().as_str()))
    {
        return Err(invalid(
            "package contains orphan or unsupported-source classic chart parts",
        ));
    }
    Ok(Graph {
        conformance,
        charts,
    })
}

/// Deterministically replace an already coherent, owned chart graph.
/// All validation completes before package mutation.
pub fn store(package: &mut OpcPackage, document_name: &PackURI, graph: &Graph) -> Result<()> {
    let current = load(package, document_name)?;
    validate_graph_value(graph)?;
    if ownership(&current) != ownership(graph) {
        return Err(invalid(
            "store cannot retarget or orphan existing chart resources",
        ));
    }
    let document = package.get_part(document_name)?;
    let (conformance, references) = document_references(document.blob())?;
    if conformance != graph.conformance
        || references
            != graph
                .charts
                .iter()
                .map(|chart| chart.document_relationship_id.clone())
                .collect::<Vec<_>>()
    {
        return Err(invalid(
            "document chart references and graph metadata differ",
        ));
    }
    for chart in &graph.charts {
        for companion in chart.styles.iter().chain(&chart.color_styles) {
            let uri = PackURI::new(&companion.part_name).map_err(Error::InvalidUri)?;
            package.add_part(Box::new(BlobPart::new(
                uri,
                companion.content_type.clone(),
                companion.data.clone(),
            )));
        }
        if let Some(workbook) = &chart.workbook {
            let uri = PackURI::new(&workbook.part_name).map_err(Error::InvalidUri)?;
            package.add_part(Box::new(BlobPart::new(
                uri,
                workbook.content_type.as_str().into(),
                workbook.data.clone(),
            )));
        }
        let chart_uri = PackURI::new(&chart.part_name).map_err(Error::InvalidUri)?;
        let mut part = BlobPart::new(
            chart_uri.clone(),
            chart.content_type.clone(),
            chart.data.clone(),
        );
        let mut relationships: Vec<(&str, &str, PackURI)> = Vec::new();
        for resource in &chart.styles {
            relationships.push((
                &resource.relationship_id,
                STYLE_REL,
                PackURI::new(&resource.part_name).map_err(Error::InvalidUri)?,
            ));
        }
        for resource in &chart.color_styles {
            relationships.push((
                &resource.relationship_id,
                COLOR_STYLE_REL,
                PackURI::new(&resource.part_name).map_err(Error::InvalidUri)?,
            ));
        }
        if let Some(workbook) = &chart.workbook {
            relationships.push((
                &workbook.relationship_id,
                graph.conformance.package_rel(),
                PackURI::new(&workbook.part_name).map_err(Error::InvalidUri)?,
            ));
        }
        relationships.sort_by(|left, right| left.0.cmp(right.0));
        for (id, kind, target) in relationships {
            part.rels_mut().add_relationship(
                kind.into(),
                target.relative_ref(chart_uri.base_uri()),
                id.to_owned(),
                false,
            );
        }
        package.add_part(Box::new(part));
    }
    let document = package.get_part_mut(document_name)?;
    let ids: Vec<_> = document
        .rels()
        .iter()
        .filter(|relationship| is_chart_rel(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect();
    for id in ids {
        document.rels_mut().remove(&id);
    }
    let mut charts: Vec<_> = graph.charts.iter().collect();
    charts.sort_by(|left, right| {
        left.document_relationship_id
            .cmp(&right.document_relationship_id)
    });
    for chart in charts {
        let target = PackURI::new(&chart.part_name).map_err(Error::InvalidUri)?;
        document.rels_mut().add_relationship(
            graph.conformance.chart_rel().into(),
            target.relative_ref(document_name.base_uri()),
            chart.document_relationship_id.clone(),
            false,
        );
    }
    Ok(())
}

fn load_companion(
    package: &OpcPackage,
    source: &dyn Part,
    relationship: &litchi_opc::Relationship,
    content_type: &str,
    root: &str,
    label: &str,
    total: &mut usize,
) -> Result<Companion> {
    let target = relationship_target(source, relationship)?;
    validate_leaf_path(&target, "/word/charts/", label)?;
    let part = package.get_part(&target)?;
    require_content_type(part, content_type, label)?;
    validate_companion_xml(part.blob(), MAX_COMPANION_XML, root, label)?;
    if part.rels().iter().next().is_some() {
        return Err(invalid(format!(
            "{label} has unsupported outbound relationships"
        )));
    }
    add_total(
        total,
        part.blob().len(),
        MAX_COMPANION_XML,
        "chart companion bytes",
    )?;
    Ok(Companion {
        relationship_id: relationship.r_id().to_owned(),
        part_name: target.as_str().to_owned(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
    })
}
