use super::model::*;
use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI, Part, Relationship};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
// Package relationship validation remains part of the canonical animation boundary.
const MAX_SLIDE_XML: usize = 64 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 64;

const P_NS: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const P_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const A_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const A_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const C_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const C_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const DGM_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DGM_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/diagram";
const R_NS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const R_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MC_NS: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const CHARTEX_URI: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";

const SLIDE_CT: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CHART_CT: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CHARTEX_CT: &str = "application/vnd.ms-office.chartex+xml";
const DIAGRAM_DATA_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
const DIAGRAM_LAYOUT_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
const DIAGRAM_STYLE_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml";
const DIAGRAM_COLORS_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";
const OLE_CT: &str = "application/vnd.openxmlformats-officedocument.oleObject";

const CHART_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
];
const CHARTEX_REL: [&str; 1] = ["http://schemas.microsoft.com/office/2014/relationships/chartEx"];
const DIAGRAM_DATA_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData",
];
const DIAGRAM_LAYOUT_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramLayout",
];
const DIAGRAM_STYLE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramQuickStyle",
];
const DIAGRAM_COLORS_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramColors",
];
const OLE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject",
];

#[derive(Debug, Clone, PartialEq, Eq)]
enum HostReference {
    Chart {
        id: String,
        extended: bool,
    },
    Diagram {
        data: String,
        layout: String,
        style: String,
        colors: String,
    },
    Ole(String),
}

#[derive(Default)]
struct FrameState {
    depth: usize,
    nv_depth: Option<usize>,
    graphic_depth: Option<usize>,
    data_depth: Option<usize>,
    data_uri: Option<String>,
    shape_id: Option<u32>,
    host: Option<HostReference>,
    enabled: bool,
}

#[derive(Default)]
struct AlternateState {
    depth: usize,
    choice_depth: Option<usize>,
    fallback_depth: Option<usize>,
    branch_enabled: bool,
    selected: bool,
    choices: usize,
    fallback_seen: bool,
}

pub fn parse_package_slide(package: &OpcPackage, slide_part_name: &PackURI) -> Result<Sequence> {
    let slide = package.get_part(slide_part_name)?;
    if slide.content_type() != SLIDE_CT {
        return invalid_relationship(
            "animation package validation requires a presentation slide part",
        );
    }
    let sequence = Sequence::parse_slide_xml(slide.blob())?;
    validate_build_relationships(package, slide, slide.blob(), &sequence)?;
    Ok(sequence)
}

fn validate_build_relationships(
    package: &OpcPackage,
    slide: &dyn Part,
    xml: &[u8],
    sequence: &Sequence,
) -> Result<()> {
    let hosts = scan_hosts(xml)?;
    for build in &sequence.graphic_builds {
        let host = hosts.get(&build.shape_id).ok_or_else(|| {
            animation_relationship_error("graphical-object build has no package-resolvable host")
        })?;
        match (&build.mode, host) {
            (GraphicBuildMode::AsOne, HostReference::Chart { id, extended })
            | (GraphicBuildMode::Chart { .. }, HostReference::Chart { id, extended }) => {
                if *extended {
                    validate_target(package, slide, id, &CHARTEX_REL, CHARTEX_CT, "/ppt/charts/")?;
                } else {
                    validate_target(package, slide, id, &CHART_REL, CHART_CT, "/ppt/charts/")?;
                }
            },
            (
                GraphicBuildMode::AsOne,
                HostReference::Diagram {
                    data,
                    layout,
                    style,
                    colors,
                },
            )
            | (
                GraphicBuildMode::Diagram { .. },
                HostReference::Diagram {
                    data,
                    layout,
                    style,
                    colors,
                },
            ) => {
                validate_diagram(package, slide, data, layout, style, colors)?;
            },
            _ => {
                return invalid_relationship(
                    "graphical-object build mode and package host subtype differ",
                );
            },
        }
    }
    for build in &sequence.diagram_builds {
        match hosts.get(&build.shape_id) {
            Some(HostReference::Ole(id)) => {
                validate_target(package, slide, id, &OLE_REL, OLE_CT, "/ppt/embeddings/")?;
            },
            _ => return invalid_relationship("diagram build has no package-resolvable OLE host"),
        }
    }
    for build in &sequence.ole_chart_builds {
        match hosts.get(&build.shape_id) {
            Some(HostReference::Ole(id)) => {
                validate_target(package, slide, id, &OLE_REL, OLE_CT, "/ppt/embeddings/")?;
            },
            _ => return invalid_relationship("OLE chart build has no package-resolvable OLE host"),
        }
    }
    Ok(())
}

fn validate_diagram(
    package: &OpcPackage,
    slide: &dyn Part,
    data: &str,
    layout: &str,
    style: &str,
    colors: &str,
) -> Result<()> {
    let ids = [data, layout, style, colors];
    if ids.iter().copied().collect::<HashSet<_>>().len() != ids.len() {
        return invalid_relationship("SmartArt relationship IDs must be distinct");
    }
    validate_target(
        package,
        slide,
        data,
        &DIAGRAM_DATA_REL,
        DIAGRAM_DATA_CT,
        "/ppt/diagrams/",
    )?;
    validate_target(
        package,
        slide,
        layout,
        &DIAGRAM_LAYOUT_REL,
        DIAGRAM_LAYOUT_CT,
        "/ppt/diagrams/",
    )?;
    validate_target(
        package,
        slide,
        style,
        &DIAGRAM_STYLE_REL,
        DIAGRAM_STYLE_CT,
        "/ppt/diagrams/",
    )?;
    validate_target(
        package,
        slide,
        colors,
        &DIAGRAM_COLORS_REL,
        DIAGRAM_COLORS_CT,
        "/ppt/diagrams/",
    )
}

fn validate_target(
    package: &OpcPackage,
    slide: &dyn Part,
    id: &str,
    relationship_types: &[&str],
    content_type: &str,
    target_prefix: &str,
) -> Result<()> {
    validate_relationship_id(id)?;
    let relationship = slide.rels().get(id).ok_or_else(|| {
        animation_relationship_error(format!("missing animation host relationship '{id}'"))
    })?;
    if relationship.is_external() {
        return invalid_relationship(format!(
            "external animation host relationship '{id}' is rejected"
        ));
    }
    if !relationship_types.contains(&relationship.reltype()) {
        return invalid_relationship(format!(
            "animation host relationship '{id}' has the wrong type"
        ));
    }
    reject_ambiguous_target(relationship)?;
    let target = relationship.target_partname().map_err(Error::Opc)?;
    if !target.as_str().starts_with(target_prefix) || target.as_str().ends_with('/') {
        return invalid_relationship(format!(
            "animation host relationship '{id}' escapes its required package directory"
        ));
    }
    let part = package.get_part(&target).map_err(|_| {
        animation_relationship_error(format!(
            "animation host relationship '{id}' targets a missing part"
        ))
    })?;
    if part.content_type() != content_type {
        return invalid_relationship(format!(
            "animation host relationship '{id}' has mismatched content type"
        ));
    }
    Ok(())
}

fn reject_ambiguous_target(relationship: &Relationship) -> Result<()> {
    let value = relationship.target_ref();
    let lower = value.to_ascii_lowercase();
    if value.contains(['?', '#', '\\'])
        || lower.contains("%2e")
        || lower.contains("%2f")
        || lower.contains("%5c")
        || value.is_empty()
    {
        return invalid_relationship("animation host relationship target is ambiguous or encoded");
    }
    Ok(())
}

fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 255
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return invalid_relationship("invalid animation host relationship ID");
    }
    Ok(())
}

fn scan_hosts(xml: &[u8]) -> Result<HashMap<u32, HostReference>> {
    if xml.len() > MAX_SLIDE_XML {
        return invalid_relationship("slide XML exceeds animation relationship scan limit");
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut frame: Option<FrameState> = None;
    let mut alternate: Option<AlternateState> = None;
    let mut hosts = HashMap::new();
    loop {
        let event = reader
            .read_event()
            .map_err(animation_relationship_xml_error)?;
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| animation_relationship_error("slide XML depth overflow"))?;
                let namespace = reader.resolver().resolve_element(element.name()).0;
                inspect_start(
                    &reader,
                    &namespace,
                    &element,
                    depth,
                    false,
                    &mut frame,
                    &mut alternate,
                )?;
            },
            Event::Empty(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| animation_relationship_error("slide XML depth overflow"))?;
                let namespace = reader.resolver().resolve_element(element.name()).0;
                inspect_start(
                    &reader,
                    &namespace,
                    &element,
                    depth,
                    true,
                    &mut frame,
                    &mut alternate,
                )?;
                inspect_end(
                    &namespace,
                    element.name(),
                    depth,
                    &mut frame,
                    &mut hosts,
                    &mut alternate,
                )?;
                depth -= 1;
            },
            Event::End(element) => {
                let namespace = reader.resolver().resolve_element(element.name()).0;
                inspect_end(
                    &namespace,
                    element.name(),
                    depth,
                    &mut frame,
                    &mut hosts,
                    &mut alternate,
                )?;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| animation_relationship_error("unbalanced slide XML"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid_relationship(
                    "DTD and processing instructions are rejected in slide XML",
                );
            },
            Event::Eof => break,
            _ => {},
        }
        nodes += 1;
        if nodes > MAX_NODES || depth > MAX_DEPTH {
            return invalid_relationship("slide XML exceeds animation relationship scan bounds");
        }
    }
    if depth != 0 || frame.is_some() || alternate.is_some() {
        return invalid_relationship("incomplete slide graphic-frame XML");
    }
    Ok(hosts)
}

fn inspect_start(
    reader: &NsReader<&[u8]>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    depth: usize,
    empty: bool,
    frame: &mut Option<FrameState>,
    alternate: &mut Option<AlternateState>,
) -> Result<()> {
    if element.attributes().with_checks(true).count() > MAX_ATTRIBUTES {
        return invalid_relationship(
            "slide element exceeds animation relationship attribute limit",
        );
    }
    if is_name(namespace, element.name(), &[MC_NS], b"AlternateContent") {
        if alternate.is_some() || empty {
            return invalid_relationship("nested or empty AlternateContent is rejected");
        }
        *alternate = Some(AlternateState {
            depth,
            ..AlternateState::default()
        });
        return Ok(());
    }
    if is_name(namespace, element.name(), &[MC_NS], b"Choice") {
        let current = alternate.as_mut().ok_or_else(|| {
            animation_relationship_error("Choice appears outside AlternateContent")
        })?;
        if depth != current.depth + 1
            || current.choice_depth.is_some()
            || current.fallback_seen
            || empty
        {
            return invalid_relationship("invalid AlternateContent Choice structure");
        }
        current.choices += 1;
        current.choice_depth = Some(depth);
        current.branch_enabled = !current.selected && requires_chartex_only(reader, element)?;
        current.selected |= current.branch_enabled;
        return Ok(());
    }
    if is_name(namespace, element.name(), &[MC_NS], b"Fallback") {
        let current = alternate.as_mut().ok_or_else(|| {
            animation_relationship_error("Fallback appears outside AlternateContent")
        })?;
        if depth != current.depth + 1
            || current.choice_depth.is_some()
            || current.fallback_seen
            || current.choices == 0
            || empty
        {
            return invalid_relationship("invalid AlternateContent Fallback structure");
        }
        current.fallback_seen = true;
        current.fallback_depth = Some(depth);
        current.branch_enabled = !current.selected;
        return Ok(());
    }
    if is_name(
        namespace,
        element.name(),
        &[P_NS, P_STRICT_NS],
        b"graphicFrame",
    ) {
        if frame.is_some() || empty {
            return invalid_relationship("nested or empty graphic frame is rejected");
        }
        let enabled = alternate.as_ref().is_none_or(|value| value.branch_enabled);
        *frame = Some(FrameState {
            depth,
            enabled,
            ..FrameState::default()
        });
        return Ok(());
    }
    let Some(current) = frame.as_mut() else {
        return Ok(());
    };
    if !current.enabled {
        return Ok(());
    }
    if depth == current.depth + 1
        && is_name(
            namespace,
            element.name(),
            &[P_NS, P_STRICT_NS],
            b"nvGraphicFramePr",
        )
    {
        if current.nv_depth.is_some() || empty {
            return invalid_relationship("graphic frame has invalid non-visual properties");
        }
        current.nv_depth = Some(depth);
    } else if current.nv_depth.is_some_and(|value| depth == value + 1)
        && is_name(namespace, element.name(), &[P_NS, P_STRICT_NS], b"cNvPr")
    {
        if current.shape_id.is_some() {
            return invalid_relationship("graphic frame has multiple direct shape IDs");
        }
        current.shape_id = Some(
            unqualified_attribute(reader, element, b"id")?
                .ok_or_else(|| animation_relationship_error("graphic frame shape ID is missing"))?
                .parse::<u32>()
                .map_err(|_| animation_relationship_error("invalid graphic frame shape ID"))?,
        );
    } else if depth == current.depth + 1
        && is_name(namespace, element.name(), &[A_NS, A_STRICT_NS], b"graphic")
    {
        if current.graphic_depth.is_some() || empty {
            return invalid_relationship("graphic frame has invalid direct graphic host");
        }
        current.graphic_depth = Some(depth);
    } else if current
        .graphic_depth
        .is_some_and(|value| depth == value + 1)
        && is_name(
            namespace,
            element.name(),
            &[A_NS, A_STRICT_NS],
            b"graphicData",
        )
    {
        if current.data_depth.is_some() || empty {
            return invalid_relationship("graphic frame has invalid direct graphic-data host");
        }
        current.data_depth = Some(depth);
        current.data_uri = unqualified_attribute(reader, element, b"uri")?;
    } else if current.data_depth.is_some_and(|value| depth == value + 1) {
        let host = if is_name(namespace, element.name(), &[C_NS, C_STRICT_NS], b"chart") {
            let extended = match current.data_uri.as_deref() {
                Some(CHARTEX_URI) => true,
                Some(value) if value.as_bytes() == C_NS || value.as_bytes() == C_STRICT_NS => false,
                _ => {
                    return invalid_relationship(
                        "chart host has an unknown or missing graphic-data URI",
                    );
                },
            };
            HostReference::Chart {
                id: required_relationship_attribute(reader, element, b"id")?,
                extended,
            }
        } else if is_name(
            namespace,
            element.name(),
            &[DGM_NS, DGM_STRICT_NS],
            b"relIds",
        ) {
            HostReference::Diagram {
                data: required_relationship_attribute(reader, element, b"dm")?,
                layout: required_relationship_attribute(reader, element, b"lo")?,
                style: required_relationship_attribute(reader, element, b"qs")?,
                colors: required_relationship_attribute(reader, element, b"cs")?,
            }
        } else if is_name(namespace, element.name(), &[P_NS, P_STRICT_NS], b"oleObj") {
            HostReference::Ole(required_relationship_attribute(reader, element, b"id")?)
        } else {
            return Ok(());
        };
        if current.host.replace(host).is_some() {
            return invalid_relationship("graphic frame has ambiguous package host references");
        }
    }
    Ok(())
}

fn inspect_end(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    depth: usize,
    frame: &mut Option<FrameState>,
    hosts: &mut HashMap<u32, HostReference>,
    alternate: &mut Option<AlternateState>,
) -> Result<()> {
    if let Some(current) = frame.as_mut() {
        if current.data_depth == Some(depth)
            && is_name(namespace, name, &[A_NS, A_STRICT_NS], b"graphicData")
        {
            current.data_depth = None;
        } else if current.graphic_depth == Some(depth)
            && is_name(namespace, name, &[A_NS, A_STRICT_NS], b"graphic")
        {
            current.graphic_depth = None;
        } else if current.nv_depth == Some(depth)
            && is_name(namespace, name, &[P_NS, P_STRICT_NS], b"nvGraphicFramePr")
        {
            current.nv_depth = None;
        } else if current.depth == depth
            && is_name(namespace, name, &[P_NS, P_STRICT_NS], b"graphicFrame")
        {
            let completed = frame.take().expect("graphic frame checked above");
            if completed.enabled
                && let (Some(shape_id), Some(host)) = (completed.shape_id, completed.host)
                && hosts.insert(shape_id, host).is_some()
            {
                return invalid_relationship("duplicate package-resolvable graphic-frame shape ID");
            }
        }
    }
    if is_name(namespace, name, &[MC_NS], b"Choice") {
        let current = alternate
            .as_mut()
            .ok_or_else(|| animation_relationship_error("unbalanced AlternateContent Choice"))?;
        if current.choice_depth != Some(depth) {
            return invalid_relationship("unbalanced AlternateContent Choice");
        }
        current.choice_depth = None;
        current.branch_enabled = false;
    } else if is_name(namespace, name, &[MC_NS], b"Fallback") {
        let current = alternate
            .as_mut()
            .ok_or_else(|| animation_relationship_error("unbalanced AlternateContent Fallback"))?;
        if current.fallback_depth != Some(depth) {
            return invalid_relationship("unbalanced AlternateContent Fallback");
        }
        current.fallback_depth = None;
        current.branch_enabled = false;
    } else if is_name(namespace, name, &[MC_NS], b"AlternateContent") {
        let current = alternate
            .take()
            .ok_or_else(|| animation_relationship_error("unbalanced AlternateContent"))?;
        if current.depth != depth
            || current.choice_depth.is_some()
            || current.fallback_depth.is_some()
            || current.choices == 0
        {
            return invalid_relationship("incomplete AlternateContent structure");
        }
    }
    Ok(())
}

fn requires_chartex_only(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<bool> {
    let value = unqualified_attribute(reader, element, b"Requires")?.ok_or_else(|| {
        animation_relationship_error("AlternateContent Choice Requires is missing")
    })?;
    let mut count = 0usize;
    for prefix in value.split_ascii_whitespace() {
        if prefix.is_empty()
            || !prefix
                .bytes()
                .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
        {
            return invalid_relationship("invalid AlternateContent Requires prefix");
        }
        count += 1;
        let qualified = format!("{prefix}:required");
        let namespace = reader
            .resolver()
            .resolve_element(QName(qualified.as_bytes()))
            .0;
        if !namespace_matches(&namespace, &[CHARTEX_URI.as_bytes()]) {
            return Ok(false);
        }
    }
    Ok(count == 1)
}

fn required_relationship_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<String> {
    namespaced_attribute(reader, element, &[R_NS, R_STRICT_NS], local)?.ok_or_else(|| {
        animation_relationship_error("animation host relationship attribute is missing")
    })
}

fn unqualified_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    relationship_attribute(reader, element, &[], local, true)
}

fn namespaced_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespaces: &[&[u8]],
    local: &[u8],
) -> Result<Option<String>> {
    relationship_attribute(reader, element, namespaces, local, false)
}

fn relationship_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespaces: &[&[u8]],
    local: &[u8],
    unqualified: bool,
) -> Result<Option<String>> {
    let mut found = None;
    for value in element.attributes().with_checks(true) {
        let value = value.map_err(animation_relationship_xml_error)?;
        let (namespace, attribute_local) = reader.resolver().resolve_attribute(value.key);
        let matches_namespace = if unqualified {
            matches!(namespace, ResolveResult::Unbound)
        } else {
            namespace_matches(&namespace, namespaces)
        };
        if matches_namespace && attribute_local.as_ref() == local {
            if found.is_some() {
                return invalid_relationship("duplicate animation host relationship attribute");
            }
            found = Some(
                value
                    .decoded_and_normalized_value(
                        quick_xml::XmlVersion::Implicit1_0,
                        reader.decoder(),
                    )
                    .map_err(animation_relationship_xml_error)?
                    .into_owned(),
            );
        }
    }
    Ok(found)
}

fn is_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    namespaces: &[&[u8]],
    local: &[u8],
) -> bool {
    namespace_matches(namespace, namespaces) && name.local_name().as_ref() == local
}

fn namespace_matches(namespace: &ResolveResult<'_>, namespaces: &[&[u8]]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if namespaces.iter().any(|expected| value.as_ref() == *expected))
}

fn animation_relationship_xml_error(error: impl std::fmt::Display) -> Error {
    Error::Invalid(format!("invalid animation relationship XML: {error}"))
}

fn animation_relationship_error(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn invalid_relationship<T>(message: impl Into<String>) -> Result<T> {
    Err(animation_relationship_error(message))
}
impl Sequence {
    /// Parse a slide timing tree and strictly validate build targets against its OPC package.
    ///
    /// Unlike the XML-only parser, this resolves chart, SmartArt, and OLE relationship IDs,
    /// requires internal existing target parts with matching relationship/content types, and
    /// never reads or executes embedded target bytes.
    pub fn from_package_slide(package: &OpcPackage, slide_part_name: &PackURI) -> Result<Self> {
        parse_package_slide(package, slide_part_name)
    }
}
