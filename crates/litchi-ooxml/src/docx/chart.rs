//! Bounded, inert classic-chart graphs owned by a DOCX main document.

use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeSet, HashSet};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CS: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const DOCUMENT_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CHART_CT: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const STYLE_CT: &str = "application/vnd.ms-office.chartstyle+xml";
const COLOR_STYLE_CT: &str = "application/vnd.ms-office.chartcolorstyle+xml";
const WORKBOOK_CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const STYLE_NS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const STYLE_REL: &str = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const COLOR_STYLE_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
const MAX_DOCUMENT_XML: usize = 32 * 1024 * 1024;
const MAX_CHART_XML: usize = 16 * 1024 * 1024;
const MAX_COMPANION_XML: usize = 4 * 1024 * 1024;
const MAX_WORKBOOK_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_BYTES: usize = 256 * 1024 * 1024;
const MAX_CHARTS: usize = 256;
const MAX_COMPANIONS: usize = 64;
const MAX_RELATIONSHIPS: usize = 130;
const MAX_NODES: usize = 200_000;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 750_000;
const MAX_ATTRIBUTE_BYTES: usize = 16 * 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DocxChartConformance {
    Transitional,
    Strict,
}

impl DocxChartConformance {
    fn w(self) -> &'static str {
        if self == Self::Strict { WS } else { W }
    }
    fn a(self) -> &'static str {
        if self == Self::Strict { AS } else { A }
    }
    fn c(self) -> &'static str {
        if self == Self::Strict { CS } else { C }
    }
    fn r(self) -> &'static str {
        if self == Self::Strict { RS } else { R }
    }
    fn chart_rel(self) -> &'static str {
        if self == Self::Strict {
            rt::STRICT_CHART
        } else {
            rt::CHART
        }
    }
    fn package_rel(self) -> &'static str {
        if self == Self::Strict {
            rt::STRICT_PACKAGE
        } else {
            rt::PACKAGE
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DocxChartEmbeddedWorkbookContentType {
    Xlsx,
}

impl DocxChartEmbeddedWorkbookContentType {
    pub fn as_str(self) -> &'static str {
        WORKBOOK_CT
    }
    fn parse(value: &str) -> Option<Self> {
        (value == WORKBOOK_CT).then_some(Self::Xlsx)
    }
    fn validates_path(self, value: &str) -> bool {
        value.to_ascii_lowercase().ends_with(".xlsx")
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocxChartCompanionResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    pub data: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocxChartEmbeddedWorkbookResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: DocxChartEmbeddedWorkbookContentType,
    pub data: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocxChartResource {
    pub document_relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    pub data: Vec<u8>,
    pub styles: Vec<DocxChartCompanionResource>,
    pub color_styles: Vec<DocxChartCompanionResource>,
    pub workbook: Option<DocxChartEmbeddedWorkbookResource>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocxChartGraph {
    pub conformance: DocxChartConformance,
    pub charts: Vec<DocxChartResource>,
}

/// Fully resolved attribute triple: `(namespace, local name, unescaped value)`.
type ResolvedAttribute = (String, String, String);
/// Fully resolved element triple: `(namespace, local name, attributes)`.
type ResolvedElement = (String, String, Vec<ResolvedAttribute>);

#[derive(Default)]
struct ChartScan {
    workbook_id: Option<String>,
    external_depth: Option<usize>,
    external_count: usize,
}

#[derive(Default)]
struct Limits {
    nodes: usize,
    attributes: usize,
    attribute_bytes: usize,
}

/// Load the complete bounded classic-chart graph owned by a DOCX main document.
pub fn load_chart_graph(package: &OpcPackage, document_name: &PackURI) -> Result<DocxChartGraph> {
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
            MAX_CHART_XML,
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
                    let content_type = DocxChartEmbeddedWorkbookContentType::parse(
                        part.content_type(),
                    )
                    .ok_or_else(|| {
                        invalid("embedded chart workbook has invalid or macro-enabled content type")
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
                    workbook = Some(DocxChartEmbeddedWorkbookResource {
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
        charts.push(DocxChartResource {
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
    Ok(DocxChartGraph {
        conformance,
        charts,
    })
}

/// Deterministically replace an already coherent, owned chart graph.
/// All validation completes before package mutation.
pub fn store_chart_graph(
    package: &mut OpcPackage,
    document_name: &PackURI,
    graph: &DocxChartGraph,
) -> Result<()> {
    let current = load_chart_graph(package, document_name)?;
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
            let uri = PackURI::new(&companion.part_name).map_err(OoxmlError::InvalidUri)?;
            package.add_part(Box::new(BlobPart::new(
                uri,
                companion.content_type.clone(),
                companion.data.clone(),
            )));
        }
        if let Some(workbook) = &chart.workbook {
            let uri = PackURI::new(&workbook.part_name).map_err(OoxmlError::InvalidUri)?;
            package.add_part(Box::new(BlobPart::new(
                uri,
                workbook.content_type.as_str().into(),
                workbook.data.clone(),
            )));
        }
        let chart_uri = PackURI::new(&chart.part_name).map_err(OoxmlError::InvalidUri)?;
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
                PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?,
            ));
        }
        for resource in &chart.color_styles {
            relationships.push((
                &resource.relationship_id,
                COLOR_STYLE_REL,
                PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?,
            ));
        }
        if let Some(workbook) = &chart.workbook {
            relationships.push((
                &workbook.relationship_id,
                graph.conformance.package_rel(),
                PackURI::new(&workbook.part_name).map_err(OoxmlError::InvalidUri)?,
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
        let target = PackURI::new(&chart.part_name).map_err(OoxmlError::InvalidUri)?;
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
) -> Result<DocxChartCompanionResource> {
    let target = relationship_target(source, relationship)?;
    validate_leaf_path(&target, "/word/charts/", label)?;
    let part = package.get_part(&target)?;
    require_content_type(part, content_type, label)?;
    validate_leaf_xml(part.blob(), MAX_COMPANION_XML, STYLE_NS, root, label)?;
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
    Ok(DocxChartCompanionResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: target.as_str().to_owned(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
    })
}

fn validate_graph_value(graph: &DocxChartGraph) -> Result<()> {
    if graph.charts.len() > MAX_CHARTS {
        return Err(limit("chart count"));
    }
    let mut total = 0usize;
    let mut document_ids = HashSet::new();
    let mut parts = HashSet::new();
    for chart in &graph.charts {
        validate_id(&chart.document_relationship_id)?;
        if !document_ids.insert(chart.document_relationship_id.as_str()) {
            return Err(invalid("document chart relationship IDs collide"));
        }
        if chart.content_type != CHART_CT {
            return Err(invalid("chart has invalid content type"));
        }
        let uri = PackURI::new(&chart.part_name).map_err(OoxmlError::InvalidUri)?;
        validate_leaf_path(&uri, "/word/charts/", "chart")?;
        if !parts.insert(chart.part_name.as_str()) {
            return Err(invalid("chart resource part names collide"));
        }
        let scan = scan_chart_xml(&chart.data, graph.conformance)?;
        add_total(&mut total, chart.data.len(), MAX_CHART_XML, "chart bytes")?;
        if chart.styles.len() > MAX_COMPANIONS || chart.color_styles.len() > MAX_COMPANIONS {
            return Err(limit("chart companion count"));
        }
        let mut ids = HashSet::new();
        for (resources, content_type, root, label) in [
            (&chart.styles, STYLE_CT, "chartStyle", "chart style"),
            (
                &chart.color_styles,
                COLOR_STYLE_CT,
                "colorStyle",
                "chart color style",
            ),
        ] {
            for resource in resources {
                validate_id(&resource.relationship_id)?;
                if !ids.insert(resource.relationship_id.as_str()) {
                    return Err(invalid("chart relationship IDs collide"));
                }
                if resource.content_type != content_type {
                    return Err(invalid(format!("{label} has invalid content type")));
                }
                let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
                validate_leaf_path(&uri, "/word/charts/", label)?;
                if !parts.insert(resource.part_name.as_str()) {
                    return Err(invalid("chart resource part names collide"));
                }
                validate_leaf_xml(&resource.data, MAX_COMPANION_XML, STYLE_NS, root, label)?;
                add_total(
                    &mut total,
                    resource.data.len(),
                    MAX_COMPANION_XML,
                    "chart companion bytes",
                )?;
            }
        }
        let workbook_id = if let Some(workbook) = &chart.workbook {
            validate_id(&workbook.relationship_id)?;
            if !ids.insert(workbook.relationship_id.as_str()) {
                return Err(invalid("chart relationship IDs collide"));
            }
            let uri = PackURI::new(&workbook.part_name).map_err(OoxmlError::InvalidUri)?;
            validate_leaf_path(&uri, "/word/embeddings/", "embedded workbook")?;
            if !workbook.content_type.validates_path(uri.as_str())
                || !parts.insert(workbook.part_name.as_str())
            {
                return Err(invalid(
                    "embedded workbook path, suffix, or ownership is invalid",
                ));
            }
            add_total(
                &mut total,
                workbook.data.len(),
                MAX_WORKBOOK_BYTES,
                "embedded workbook bytes",
            )?;
            Some(workbook.relationship_id.as_str())
        } else {
            None
        };
        if scan.workbook_id.as_deref() != workbook_id {
            return Err(invalid(
                "chart externalData and embedded workbook metadata differ",
            ));
        }
    }
    Ok(())
}

fn document_references(xml: &[u8]) -> Result<(DocxChartConformance, Vec<String>)> {
    for conformance in [
        DocxChartConformance::Transitional,
        DocxChartConformance::Strict,
    ] {
        if let Ok(value) = scan_document_xml(xml, conformance) {
            return Ok((conformance, value));
        }
    }
    Err(invalid("invalid DOCX document root or chart anchors"))
}

fn scan_document_xml(xml: &[u8], conformance: DocxChartConformance) -> Result<Vec<String>> {
    if xml.len() > MAX_DOCUMENT_XML {
        return Err(limit("document XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_DOCUMENT_XML {
        return Err(limit("processed document XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut limits = Limits::default();
    let mut root = false;
    let mut frames: Vec<(usize, usize)> = Vec::new();
    let mut references = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth += 1;
                structure(&mut limits, depth)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != conformance.w() || local != "document" {
                        return Err(invalid("invalid document root or namespace"));
                    }
                    root = true;
                }
                if ns == conformance.a() && local == "graphicData" {
                    if attr(&attrs, "", "uri") == Some(conformance.c()) {
                        frames.push((depth, 0));
                    }
                } else if ns == conformance.c() && local == "chart" {
                    let Some(frame) = frames.last_mut() else {
                        return Err(invalid("chart element is outside chart graphicData"));
                    };
                    frame.1 += 1;
                    let id = required_rel_id(&attrs, conformance)?;
                    references.push(id.to_owned());
                }
            },
            Event::Empty(element) => {
                structure(&mut limits, depth + 1)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != conformance.w() || local != "document" {
                        return Err(invalid("invalid document root or namespace"));
                    }
                    root = true;
                }
                if ns == conformance.a()
                    && local == "graphicData"
                    && attr(&attrs, "", "uri") == Some(conformance.c())
                {
                    return Err(invalid("chart graphicData lacks chart child"));
                }
                if ns == conformance.c() && local == "chart" {
                    let Some(frame) = frames.last_mut() else {
                        return Err(invalid("chart element is outside chart graphicData"));
                    };
                    frame.1 += 1;
                    references.push(required_rel_id(&attrs, conformance)?.to_owned());
                }
            },
            Event::End(_) => {
                if let Some((_, count)) = frames.pop_if(|frame| frame.0 == depth)
                    && count != 1
                {
                    return Err(invalid(
                        "chart graphicData must contain exactly one chart reference",
                    ));
                }
                if depth == 0 {
                    return Err(invalid("unexpected document XML closing element"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root || depth != 0 || !frames.is_empty() {
        return Err(invalid("missing or unterminated document root"));
    }
    Ok(references)
}

fn scan_chart_xml(xml: &[u8], conformance: DocxChartConformance) -> Result<ChartScan> {
    if xml.len() > MAX_CHART_XML {
        return Err(limit("chart XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_CHART_XML {
        return Err(limit("processed chart XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut limits = Limits::default();
    let mut root = false;
    let mut scan = ChartScan::default();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth += 1;
                structure(&mut limits, depth)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != conformance.c() || local != "chartSpace" {
                        return Err(invalid("invalid chartSpace root or namespace"));
                    }
                    if attr(&attrs, "", "fallbackImg").is_some() {
                        return Err(invalid("chart fallback image relationship is unsupported"));
                    }
                    root = true;
                }
                inspect_chart_element(&ns, &local, &attrs, conformance, depth, &mut scan)?;
            },
            Event::Empty(element) => {
                structure(&mut limits, depth + 1)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != conformance.c() || local != "chartSpace" {
                        return Err(invalid("invalid chartSpace root or namespace"));
                    }
                    if attr(&attrs, "", "fallbackImg").is_some() {
                        return Err(invalid("chart fallback image relationship is unsupported"));
                    }
                    root = true;
                }
                inspect_chart_element(&ns, &local, &attrs, conformance, depth + 1, &mut scan)?;
                if ns == conformance.c() && local == "externalData" {
                    scan.external_depth = None;
                }
            },
            Event::End(_) => {
                if scan.external_depth == Some(depth) {
                    scan.external_depth = None;
                }
                if depth == 0 {
                    return Err(invalid("unexpected chart XML closing element"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root || depth != 0 {
        return Err(invalid("missing or unterminated chartSpace root"));
    }
    Ok(scan)
}

fn inspect_chart_element(
    ns: &str,
    local: &str,
    attrs: &[ResolvedAttribute],
    conformance: DocxChartConformance,
    depth: usize,
    scan: &mut ChartScan,
) -> Result<()> {
    if ns == conformance.c() && local == "externalData" {
        scan.external_count += 1;
        if scan.external_count > 1 {
            return Err(invalid("chart has multiple externalData elements"));
        }
        scan.external_depth = Some(depth);
        scan.workbook_id = Some(required_rel_id(attrs, conformance)?.to_owned());
    } else if ns == conformance.c() && local == "autoUpdate" {
        if scan.external_depth.is_none() {
            return Err(invalid("autoUpdate is outside externalData"));
        }
        match attr(attrs, "", "val") {
            Some("0" | "false") => {},
            _ => return Err(invalid("automatic chart data updates are rejected")),
        }
    }
    for (namespace, name, _) in attrs {
        if matches!(namespace.as_str(), R | RS)
            && !(ns == conformance.c()
                && local == "externalData"
                && namespace == conformance.r()
                && name == "id")
        {
            return Err(invalid(
                "chart XML contains an unsupported relationship reference",
            ));
        }
    }
    Ok(())
}

fn validate_leaf_xml(
    xml: &[u8],
    max: usize,
    namespace: &str,
    root_name: &str,
    label: &str,
) -> Result<()> {
    if xml.len() > max {
        return Err(limit("companion XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > max {
        return Err(limit("processed companion XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut limits = Limits::default();
    let mut root = false;
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth += 1;
                structure(&mut limits, depth)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != namespace || local != root_name {
                        return Err(invalid(format!("invalid {label} root or namespace")));
                    }
                    root = true;
                }
                if attrs.iter().any(|(ns, _, _)| matches!(ns.as_str(), R | RS)) {
                    return Err(invalid(format!(
                        "{label} contains an unsupported relationship reference"
                    )));
                }
            },
            Event::Empty(element) => {
                structure(&mut limits, depth + 1)?;
                let (ns, local, attrs) = element_info(&reader, &element, &mut limits)?;
                if !root {
                    if ns != namespace || local != root_name {
                        return Err(invalid(format!("invalid {label} root or namespace")));
                    }
                    root = true;
                }
                if attrs.iter().any(|(ns, _, _)| matches!(ns.as_str(), R | RS)) {
                    return Err(invalid(format!(
                        "{label} contains an unsupported relationship reference"
                    )));
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid(format!("unexpected {label} closing element")));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root || depth != 0 {
        return Err(invalid(format!("missing or unterminated {label} root")));
    }
    Ok(())
}

fn element_info(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: &mut Limits,
) -> Result<ResolvedElement> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut values = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        limits.attributes += 1;
        if limits.attributes > MAX_ATTRIBUTES {
            return Err(limit("XML attribute count"));
        }
        let (namespace, name) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(name.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let raw_value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(raw_value)
            .map_err(xml_error)?
            .into_owned();
        limits.attribute_bytes = limits
            .attribute_bytes
            .checked_add(namespace.len() + name.len() + value.len())
            .ok_or_else(|| limit("XML attribute bytes"))?;
        if limits.attribute_bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("XML attribute bytes"));
        }
        if values
            .iter()
            .any(|(ns, n, _): &ResolvedAttribute| ns == &namespace && n == &name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        values.push((namespace, name, value));
    }
    Ok((namespace, local, values))
}

fn required_rel_id(attrs: &[ResolvedAttribute], conformance: DocxChartConformance) -> Result<&str> {
    let value = attr(attrs, conformance.r(), "id")
        .ok_or_else(|| invalid("chart reference lacks relationship ID"))?;
    if attrs
        .iter()
        .any(|(namespace, name, _)| !(namespace == conformance.r() && name == "id"))
    {
        return Err(invalid("chart reference has unsupported attributes"));
    }
    validate_id(value)?;
    Ok(value)
}
fn attr<'a>(attrs: &'a [ResolvedAttribute], namespace: &str, name: &str) -> Option<&'a str> {
    attrs
        .iter()
        .find(|(ns, n, _)| ns == namespace && n == name)
        .map(|(_, _, value)| value.as_str())
}
fn structure(limits: &mut Limits, depth: usize) -> Result<()> {
    limits.nodes += 1;
    if limits.nodes > MAX_NODES || depth > MAX_DEPTH {
        return Err(limit("XML structure"));
    }
    Ok(())
}
fn ownership(graph: &DocxChartGraph) -> BTreeSet<String> {
    graph
        .charts
        .iter()
        .flat_map(|chart| {
            std::iter::once(chart.part_name.clone())
                .chain(
                    chart
                        .styles
                        .iter()
                        .map(|resource| resource.part_name.clone()),
                )
                .chain(
                    chart
                        .color_styles
                        .iter()
                        .map(|resource| resource.part_name.clone()),
                )
                .chain(
                    chart
                        .workbook
                        .iter()
                        .map(|resource| resource.part_name.clone()),
                )
        })
        .collect()
}
fn relationship_target(
    part: &dyn Part,
    relationship: &litchi_opc::Relationship,
) -> Result<PackURI> {
    if relationship.is_external() {
        return Err(invalid("external relationship is rejected"));
    }
    PackURI::from_rel_ref(part.partname().base_uri(), relationship.target_ref())
        .map_err(OoxmlError::InvalidFormat)
}
fn validate_leaf_path(uri: &PackURI, prefix: &str, label: &str) -> Result<()> {
    let Some(rest) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(format!("{label} is outside {prefix}")));
    };
    if rest.is_empty()
        || rest.contains('/')
        || !rest
            .to_ascii_lowercase()
            .ends_with(if label == "embedded workbook" {
                ".xlsx"
            } else {
                ".xml"
            })
    {
        return Err(invalid(format!("invalid {label} path or suffix")));
    }
    Ok(())
}
fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(invalid(format!("{label} has invalid content type")))
    }
}
fn is_chart_rel(value: &str) -> bool {
    matches!(value, rt::CHART | rt::STRICT_CHART)
}
fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID is empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid("invalid relationship ID"))
    } else {
        Ok(())
    }
}
fn add_total(total: &mut usize, size: usize, individual: usize, label: &str) -> Result<()> {
    if size > individual {
        return Err(limit(label));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("aggregate bytes"))?;
    if *total > MAX_TOTAL_BYTES {
        return Err(limit("aggregate bytes"));
    }
    Ok(())
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(label: &str) -> OoxmlError {
    invalid(format!("DOCX chart {label} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    const POI: &[u8] = include_bytes!("../../../../test-data/poi/test-data/document/61745.docx");
    const LO_INTERNAL: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/layout-in-cell-2.docx"
    );
    const LO_EXTERNAL: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/oox/qa/unit/data/chart-data-label-char-color.docx"
    );
    fn document() -> PackURI {
        PackURI::new("/word/document.xml").unwrap()
    }
    #[test]
    fn poi_and_libreoffice_internal_charts_round_trip_deterministically() {
        for (bytes, count) in [(POI, 2usize), (LO_INTERNAL, 1usize)] {
            let mut package = OpcPackage::from_bytes(bytes).unwrap();
            let name = document();
            let graph = load_chart_graph(&package, &name).unwrap();
            assert_eq!(graph.charts.len(), count);
            assert!(graph.charts.iter().all(|chart| chart.workbook.is_some()));
            store_chart_graph(&mut package, &name, &graph).unwrap();
            assert_eq!(load_chart_graph(&package, &name).unwrap(), graph);
            store_chart_graph(&mut package, &name, &graph).unwrap();
            assert_eq!(load_chart_graph(&package, &name).unwrap(), graph);
        }
    }
    fn synthetic(conformance: DocxChartConformance) -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let name = document();
        let mut document_part=BlobPart::new(name.clone(),DOCUMENT_CT.into(),format!("<w:document xmlns:w=\"{}\" xmlns:a=\"{}\" xmlns:c=\"{}\" xmlns:r=\"{}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><w:body><mc:AlternateContent><mc:Choice Requires=\"u\"><u:active/></mc:Choice><mc:Fallback><w:p><w:r><w:drawing><a:graphic><a:graphicData uri=\"{}\"><c:chart r:id=\"rIdChart\"/></a:graphicData></a:graphic></w:drawing></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>",conformance.w(),conformance.a(),conformance.c(),conformance.r(),conformance.c()).into_bytes());
        document_part.rels_mut().add_relationship(
            conformance.chart_rel().into(),
            "charts/chart1.xml".into(),
            "rIdChart".into(),
            false,
        );
        package.add_part(Box::new(document_part));
        let chart_uri = PackURI::new("/word/charts/chart1.xml").unwrap();
        let mut chart=BlobPart::new(chart_uri,CHART_CT.into(),format!("<c:chartSpace xmlns:c=\"{}\" xmlns:r=\"{}\"><c:chart/><c:externalData r:id=\"rIdWorkbook\"><c:autoUpdate val=\"0\"/></c:externalData></c:chartSpace>",conformance.c(),conformance.r()).into_bytes());
        chart.rels_mut().add_relationship(
            STYLE_REL.into(),
            "style1.xml".into(),
            "rIdStyle".into(),
            false,
        );
        chart.rels_mut().add_relationship(
            COLOR_STYLE_REL.into(),
            "colors1.xml".into(),
            "rIdColors".into(),
            false,
        );
        chart.rels_mut().add_relationship(
            conformance.package_rel().into(),
            "../embeddings/data1.xlsx".into(),
            "rIdWorkbook".into(),
            false,
        );
        package.add_part(Box::new(chart));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/charts/style1.xml").unwrap(),
            STYLE_CT.into(),
            format!("<cs:chartStyle xmlns:cs=\"{STYLE_NS}\"/>").into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/charts/colors1.xml").unwrap(),
            COLOR_STYLE_CT.into(),
            format!("<cs:colorStyle xmlns:cs=\"{STYLE_NS}\"/>").into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/embeddings/data1.xlsx").unwrap(),
            WORKBOOK_CT.into(),
            b"PK opaque workbook".to_vec(),
        )));
        (package, name)
    }
    #[test]
    fn strict_mce_graph_round_trips_without_opening_workbook() {
        let (mut package, name) = synthetic(DocxChartConformance::Strict);
        let graph = load_chart_graph(&package, &name).unwrap();
        assert_eq!(
            graph.charts[0].workbook.as_ref().unwrap().data,
            b"PK opaque workbook"
        );
        store_chart_graph(&mut package, &name, &graph).unwrap();
        assert_eq!(load_chart_graph(&package, &name).unwrap(), graph);
    }
    #[test]
    fn rejects_external_ole_malformed_orphan_unsupported_and_caps_before_mutation() {
        let package = OpcPackage::from_bytes(LO_EXTERNAL).unwrap();
        assert!(load_chart_graph(&package, &document()).is_err());
        let (mut package, name) = synthetic(DocxChartConformance::Transitional);
        package
            .get_part_mut(&PackURI::new("/word/charts/chart1.xml").unwrap())
            .unwrap()
            .set_blob(format!("<c:wrong xmlns:c=\"{C}\"/>").into_bytes());
        assert!(load_chart_graph(&package, &name).is_err());
        let (mut package, name) = synthetic(DocxChartConformance::Transitional);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/charts/orphan.xml").unwrap(),
            CHART_CT.into(),
            format!("<c:chartSpace xmlns:c=\"{C}\"><c:chart/></c:chartSpace>").into_bytes(),
        )));
        assert!(load_chart_graph(&package, &name).is_err());
        let (mut package, name) = synthetic(DocxChartConformance::Transitional);
        package
            .get_part_mut(&PackURI::new("/word/charts/chart1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::IMAGE.into(),
                "../media/image1.png".into(),
                "rIdImage".into(),
                false,
            );
        assert!(load_chart_graph(&package, &name).is_err());
        let (mut package, name) = synthetic(DocxChartConformance::Transitional);
        let mut graph = load_chart_graph(&package, &name).unwrap();
        graph.charts[0].workbook.as_mut().unwrap().data = vec![0; MAX_WORKBOOK_BYTES + 1];
        let before = package.get_part(&name).unwrap().blob().to_vec();
        assert!(store_chart_graph(&mut package, &name, &graph).is_err());
        assert_eq!(package.get_part(&name).unwrap().blob(), before);
    }
}
