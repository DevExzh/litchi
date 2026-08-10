//! Bounded worksheet DrawingML picture/chart transfer planning.

use std::collections::{BTreeSet, HashSet};

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part, Relationship, TargetMode};
use litchi_sheet::{Cell as Address, Rect};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;

use super::{GraphAction, GraphChange, Workbook, Worksheet, allocation, invalid};
use crate::error::{Error, Result};

const XDR: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_DRAWING_BYTES: usize = 32 * 1024 * 1024;
const MAX_LEAF_BYTES: usize = 32 * 1024 * 1024;
const MAX_ANCHORS: usize = 100_000;
const MAX_REFERENCES: usize = 4_096;
const MAX_URI_ATTEMPTS: u32 = 10_000;

#[derive(Debug)]
pub(super) struct Plan {
    pub(super) source_position: usize,
    pub(super) source_name: Box<str>,
    pub(super) target_name: Box<str>,
    pub(super) target_relationship_id: String,
    pub(super) anchors: usize,
    pub(super) graph: Vec<GraphChange>,
}

#[derive(Debug)]
struct Projection {
    xml: Vec<u8>,
    relationship_ids: Vec<String>,
    anchors: usize,
}

#[derive(Debug)]
struct AnchorSpan {
    start: usize,
    end: usize,
}

#[derive(Debug)]
struct Anchor {
    from_row: u32,
    from_column: u32,
    relationship_ids: Vec<String>,
}

pub(super) fn plan(
    workbook: &Workbook,
    source: &Worksheet,
    target: &Worksheet,
    source_range: Rect,
    target_start: Address,
) -> Result<Option<Plan>> {
    let source_part = workbook.inner.package.get_part(&source.data.part_uri)?;
    let drawings = source_part
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING))
        .collect::<Vec<_>>();
    if drawings.is_empty() {
        return Ok(None);
    }
    if drawings.len() != 1 {
        return Err(unsupported(
            "copying worksheets with multiple drawing parts",
        ));
    }
    let source_relationship = drawings[0];
    if source_relationship.is_external() {
        return Err(invalid("worksheet drawing relationship cannot be external"));
    }
    let source_reference = worksheet_drawing_reference(source_part.blob())?
        .ok_or_else(|| invalid("worksheet drawing relationship has no direct drawing reference"))?;
    if source_reference != source_relationship.r_id() {
        return Err(invalid(
            "worksheet drawing reference does not match its package relationship",
        ));
    }

    let target_part = workbook.inner.package.get_part(&target.data.part_uri)?;
    if target_part
        .rels()
        .iter()
        .any(|relationship| matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING))
        || worksheet_drawing_reference(target_part.blob())?.is_some()
    {
        return Err(unsupported("merging into an existing worksheet drawing"));
    }

    let drawing_uri = source_relationship.target_partname()?;
    let drawing_part = workbook.inner.package.get_part(&drawing_uri)?;
    if drawing_part.content_type() != ct::OFC_DRAWING {
        return Err(invalid(format!(
            "worksheet drawing part has content type '{}', expected '{}'",
            drawing_part.content_type(),
            ct::OFC_DRAWING
        )));
    }
    let row_delta = i64::from(target_start.row().get())
        .checked_sub(i64::from(source_range.start().row().get()))
        .ok_or_else(|| invalid("drawing row translation overflow"))?;
    let column_delta = i64::from(target_start.column().get())
        .checked_sub(i64::from(source_range.start().column().get()))
        .ok_or_else(|| invalid("drawing column translation overflow"))?;
    let Some(projection) =
        project_drawing(drawing_part.blob(), source_range, row_delta, column_delta)?
    else {
        return Ok(None);
    };

    let mut reserved = workbook
        .inner
        .package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<BTreeSet<_>>();
    let target_drawing_uri = allocate_uri(&drawing_uri, &mut reserved)?;
    let target_relationship_id = allocate_relationship_id(target_part)?;
    let mut target_drawing = BlobPart::new(
        target_drawing_uri.clone(),
        drawing_part.content_type().to_owned(),
        projection.xml,
    );
    let mut internal = Vec::new();
    internal
        .try_reserve_exact(projection.relationship_ids.len())
        .map_err(|source| allocation("drawing dependency plan", source))?;
    for relationship_id in &projection.relationship_ids {
        let relationship = drawing_part.rels().get(relationship_id).ok_or_else(|| {
            invalid(format!(
                "drawing anchor references missing relationship '{relationship_id}'"
            ))
        })?;
        if relationship.is_external() {
            if matches!(relationship.reltype(), rt::CHART | rt::STRICT_CHART) {
                return Err(invalid("drawing chart relationship cannot be external"));
            }
            target_drawing.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                TargetMode::External,
            )?;
            continue;
        }
        let source_leaf_uri = relationship.target_partname()?;
        let source_leaf = workbook.inner.package.get_part(&source_leaf_uri)?;
        let target_leaf_uri = allocate_uri(&source_leaf_uri, &mut reserved)?;
        let mut target_leaf = BlobPart::new_shared(
            target_leaf_uri.clone(),
            source_leaf.content_type().to_owned(),
            source_leaf.blob_arc(),
        );
        let descendants = prepare_dependency(
            workbook,
            relationship.reltype(),
            source_leaf,
            &target_leaf_uri,
            &mut target_leaf,
            &mut reserved,
        )?;
        if internal
            .len()
            .checked_add(1usize.saturating_add(descendants.len()))
            .is_none_or(|count| count > MAX_REFERENCES)
        {
            return Err(invalid("drawing dependency graph part limit exceeded"));
        }
        let cloned_relationship = Relationship::new_with_mode(
            relationship.r_id().to_owned(),
            relationship.reltype().to_owned(),
            target_leaf_uri.relative_ref(target_drawing_uri.base_uri()),
            target_drawing_uri.base_uri().to_owned(),
            TargetMode::Internal,
        );
        internal.push(GraphChange {
            action: GraphAction::Add,
            source: target_drawing_uri.clone(),
            relationship: cloned_relationship,
            part: Box::new(target_leaf),
        });
        internal.extend(descendants);
    }
    let worksheet_relationship = Relationship::new_with_mode(
        target_relationship_id.clone(),
        source_relationship.reltype().to_owned(),
        target_drawing_uri.relative_ref(target.data.part_uri.base_uri()),
        target.data.part_uri.base_uri().to_owned(),
        TargetMode::Internal,
    );
    let mut graph = Vec::new();
    graph
        .try_reserve_exact(1usize.saturating_add(internal.len()))
        .map_err(|source| allocation("drawing graph plan", source))?;
    graph.push(GraphChange {
        action: GraphAction::Add,
        source: target.data.part_uri.clone(),
        relationship: worksheet_relationship,
        part: Box::new(target_drawing),
    });
    graph.extend(internal);
    Ok(Some(Plan {
        source_position: source.position(),
        source_name: source.name().into(),
        target_name: target.name().into(),
        target_relationship_id,
        anchors: projection.anchors,
        graph,
    }))
}

fn prepare_dependency(
    workbook: &Workbook,
    relationship_type: &str,
    source: &dyn Part,
    target_uri: &PackURI,
    target: &mut BlobPart,
    reserved: &mut BTreeSet<String>,
) -> Result<Vec<GraphChange>> {
    if source.blob().len() > MAX_LEAF_BYTES {
        return Err(invalid("drawing dependency exceeds the size limit"));
    }
    match relationship_type {
        rt::IMAGE | rt::STRICT_IMAGE => {
            if !source.rels().is_empty() {
                return Err(unsupported(
                    "copying image parts with outbound relationships",
                ));
            }
            validate_image_leaf(source)?;
            Ok(Vec::new())
        },
        rt::CHART | rt::STRICT_CHART => {
            if source.content_type() != ct::DML_CHART {
                return Err(invalid(format!(
                    "drawing chart part has content type '{}', expected '{}'",
                    source.content_type(),
                    ct::DML_CHART
                )));
            }
            if !source.partname().as_str().starts_with("/xl/charts/")
                || !source.partname().as_str().ends_with(".xml")
            {
                return Err(invalid(
                    "drawing chart target is outside /xl/charts or lacks .xml suffix",
                ));
            }
            let anchor = crate::chart::Anchor::new(0, 0, 1, 1);
            let chart = crate::chart::read(source.blob(), anchor)?;
            clone_chart_relationships(workbook, source, target_uri, target, reserved, &chart)
        },
        _ => Err(unsupported(
            "copying drawing dependencies other than images and classic charts",
        )),
    }
}

fn clone_chart_relationships(
    workbook: &Workbook,
    source: &dyn Part,
    target_uri: &PackURI,
    target: &mut BlobPart,
    reserved: &mut BTreeSet<String>,
    chart: &crate::chart::Chart,
) -> Result<Vec<GraphChange>> {
    let mut fragment_ids =
        litchi_spreadsheet_drawing::chart::relationship::fragment_ids(&chart.chart)?;
    let external_data_id = chart
        .chart
        .external_data
        .as_ref()
        .and_then(|value| value.relationship_id.as_deref());
    let user_shapes_id = chart
        .chart
        .user_shapes
        .as_ref()
        .and_then(|value| value.relationship_id.as_deref());
    for id in [external_data_id, user_shapes_id].into_iter().flatten() {
        if !fragment_ids.insert(id.to_owned()) {
            return Err(invalid("chart relationship roles use the same ID"));
        }
    }
    let actual = source
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<HashSet<_>>();
    if actual != fragment_ids {
        return Err(invalid(
            "chart XML relationship references do not close over package relationships",
        ));
    }
    if actual.len() > MAX_REFERENCES {
        return Err(invalid("chart relationship limit exceeded"));
    }
    let mut relationships = source.rels().iter().collect::<Vec<_>>();
    relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    let mut graph = Vec::new();
    graph
        .try_reserve_exact(relationships.len())
        .map_err(|source| allocation("chart dependency graph", source))?;
    for relationship in relationships {
        validate_chart_relationship_role(relationship, external_data_id, user_shapes_id)?;
        if relationship.is_external() {
            target.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                TargetMode::External,
            )?;
            continue;
        }
        let source_child_uri = relationship.target_partname()?;
        let source_child = workbook.inner.package.get_part(&source_child_uri)?;
        validate_chart_child(relationship, source_child, user_shapes_id)?;
        let target_child_uri = allocate_uri(&source_child_uri, reserved)?;
        let target_child = BlobPart::new_shared(
            target_child_uri.clone(),
            source_child.content_type().to_owned(),
            source_child.blob_arc(),
        );
        let cloned_relationship = Relationship::new_with_mode(
            relationship.r_id().to_owned(),
            relationship.reltype().to_owned(),
            target_child_uri.relative_ref(target_uri.base_uri()),
            target_uri.base_uri().to_owned(),
            TargetMode::Internal,
        );
        graph.push(GraphChange {
            action: GraphAction::Add,
            source: target_uri.clone(),
            relationship: cloned_relationship,
            part: Box::new(target_child),
        });
    }
    Ok(graph)
}

fn validate_chart_relationship_role(
    relationship: &Relationship,
    external_data_id: Option<&str>,
    user_shapes_id: Option<&str>,
) -> Result<()> {
    if external_data_id == Some(relationship.r_id()) {
        if !litchi_spreadsheet_drawing::chart::relationship::is_external_data_type(
            relationship.reltype(),
        ) {
            return Err(invalid("chart external-data relationship has invalid type"));
        }
        return Ok(());
    }
    if user_shapes_id == Some(relationship.r_id()) {
        if relationship.is_external()
            || !litchi_spreadsheet_drawing::chart::relationship::is_user_shapes_type(
                relationship.reltype(),
            )
        {
            return Err(invalid(
                "chart user-shapes relationship has invalid type or mode",
            ));
        }
        return Ok(());
    }
    if relationship.is_external()
        && matches!(
            relationship.reltype(),
            rt::HYPERLINK | rt::STRICT_HYPERLINK | rt::IMAGE | rt::STRICT_IMAGE
        )
    {
        return Ok(());
    }
    if !relationship.is_external() && matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
    {
        return Ok(());
    }
    Err(unsupported(
        "copying unsupported classic-chart relationships",
    ))
}

fn validate_chart_child(
    relationship: &Relationship,
    child: &dyn Part,
    user_shapes_id: Option<&str>,
) -> Result<()> {
    if !child.rels().is_empty() {
        return Err(unsupported(
            "copying classic-chart dependencies with outbound relationships",
        ));
    }
    if child.blob().len() > MAX_LEAF_BYTES {
        return Err(invalid("chart dependency exceeds the size limit"));
    }
    if user_shapes_id == Some(relationship.r_id()) {
        if child.content_type() != ct::DML_CHARTSHAPES
            || !child.partname().as_str().starts_with("/xl/drawings/")
            || !child.partname().as_str().ends_with(".xml")
        {
            return Err(invalid(
                "chart user-shapes target has invalid path or content type",
            ));
        }
        if !litchi_spreadsheet_drawing::chart::relationship::user_shapes_ids(child.blob())?
            .is_empty()
        {
            return Err(unsupported(
                "copying chart user-shapes with referenced dependencies",
            ));
        }
        return Ok(());
    }
    if matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE) {
        return validate_image_leaf(child);
    }
    let expected = litchi_spreadsheet_drawing::chart::relationship::external_data_content_type(
        relationship.reltype(),
    )
    .ok_or_else(|| unsupported("copying unsupported embedded chart data"))?;
    if !child.partname().as_str().starts_with("/xl/embeddings/") {
        return Err(invalid(
            "embedded chart data target is outside /xl/embeddings",
        ));
    }
    if child.content_type() != expected {
        return Err(invalid(format!(
            "embedded chart data has content type '{}', expected '{expected}'",
            child.content_type()
        )));
    }
    Ok(())
}

fn validate_image_leaf(part: &dyn Part) -> Result<()> {
    let name = part.partname().as_str();
    if !name.starts_with("/xl/media/") {
        return Err(invalid("drawing image target is outside /xl/media"));
    }
    let lower = name.to_ascii_lowercase();
    let suffix_matches = match part.content_type() {
        "image/bmp" => lower.ends_with(".bmp"),
        "image/gif" => lower.ends_with(".gif"),
        "image/png" => lower.ends_with(".png"),
        "image/tif" => lower.ends_with(".tif"),
        "image/tiff" => lower.ends_with(".tiff"),
        "image/x-icon" => lower.ends_with(".ico"),
        "image/x-pcx" => lower.ends_with(".pcx"),
        "image/jpeg" => lower.ends_with(".jpg") || lower.ends_with(".jpeg"),
        "image/jp2" => lower.ends_with(".jp2"),
        "image/x-emf" => lower.ends_with(".emf"),
        "image/x-wmf" => lower.ends_with(".wmf"),
        "image/svg+xml" => lower.ends_with(".svg"),
        _ => false,
    };
    if !suffix_matches {
        return Err(invalid(format!(
            "drawing image path '{}' does not match supported content type '{}'",
            part.partname(),
            part.content_type(),
        )));
    }
    Ok(())
}

pub(super) fn attach_worksheet(xml: &[u8], relationship_id: &str) -> Result<Vec<u8>> {
    let layout = worksheet_layout(xml)?;
    if layout.drawing_reference.is_some() {
        return Err(unsupported("merging into an existing worksheet drawing"));
    }
    if layout.alternate_content {
        return Err(unsupported(
            "inserting a drawing beside worksheet markup-compatibility branches",
        ));
    }
    let prefix = layout
        .root_name
        .split_once(':')
        .map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
    let fragment = format!(
        "<{prefix}drawing xmlns:r=\"{}\" r:id=\"{}\"/>",
        String::from_utf8_lossy(layout.relationship_namespace),
        litchi_core::xml::escape_xml(relationship_id)
    );
    let capacity = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| invalid("worksheet drawing insertion size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("worksheet drawing insertion", source))?;
    output.extend_from_slice(&xml[..layout.insertion]);
    output.extend_from_slice(fragment.as_bytes());
    output.extend_from_slice(&xml[layout.insertion..]);
    crate::raw::compact::changed(&output, "compact worksheet drawing insertion")
}

fn project_drawing(
    xml: &[u8],
    source_range: Rect,
    row_delta: i64,
    column_delta: i64,
) -> Result<Option<Projection>> {
    if xml.len() > MAX_DRAWING_BYTES {
        return Err(invalid("worksheet drawing XML exceeds the size limit"));
    }
    let layout = drawing_layout(xml)?;
    let mut selected = Vec::new();
    let mut relationship_ids = BTreeSet::new();
    for span in layout.anchors {
        let anchor = parse_anchor(xml, &span)?;
        let address = Address::at(anchor.from_row, anchor.from_column)?;
        if !source_range.contains(address) {
            continue;
        }
        for relationship_id in anchor.relationship_ids {
            if relationship_ids.len() >= MAX_REFERENCES {
                return Err(invalid("drawing relationship reference limit exceeded"));
            }
            relationship_ids.insert(relationship_id);
        }
        selected.push(translate_anchor(xml, &span, row_delta, column_delta)?);
    }
    if selected.is_empty() {
        return Ok(None);
    }
    let selected_bytes = selected.iter().map(Vec::len).sum::<usize>();
    let capacity = layout
        .root_open_end
        .checked_add(selected_bytes)
        .and_then(|size| size.checked_add(xml.len().saturating_sub(layout.root_close_start)))
        .ok_or_else(|| invalid("projected drawing size overflow"))?;
    if capacity > MAX_DRAWING_BYTES {
        return Err(invalid("projected drawing XML exceeds the size limit"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("projected worksheet drawing", source))?;
    output.extend_from_slice(&xml[..layout.root_open_end]);
    for anchor in &selected {
        output.extend_from_slice(anchor);
    }
    output.extend_from_slice(&xml[layout.root_close_start..]);
    let output = crate::raw::compact::changed(&output, "compact transferred worksheet drawing")?;
    Ok(Some(Projection {
        xml: output,
        relationship_ids: relationship_ids.into_iter().collect(),
        anchors: selected.len(),
    }))
}

struct DrawingLayout {
    root_open_end: usize,
    root_close_start: usize,
    anchors: Vec<AnchorSpan>,
}

fn drawing_layout(xml: &[u8]) -> Result<DrawingLayout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_open_end = None;
    let mut root_close_start = None;
    let mut open_anchor = None;
    let mut anchors = Vec::new();
    loop {
        let start = position(&reader)?;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if !xdr(&namespace) || element.local_name().as_ref() != b"wsDr" {
                        return Err(invalid("drawing transfer requires one xdr:wsDr root"));
                    }
                    root_open_end = Some(end);
                } else if depth == 1 && xdr(&namespace) {
                    match element.local_name().as_ref() {
                        b"twoCellAnchor" | b"oneCellAnchor" => open_anchor = Some(start),
                        b"absoluteAnchor" => {},
                        _ => return Err(unsupported("copying unknown drawing root children")),
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("drawing XML nesting overflow"))?;
            },
            Event::Empty(element) if depth == 1 && xdr(&namespace) => {
                if matches!(
                    element.local_name().as_ref(),
                    b"twoCellAnchor" | b"oneCellAnchor"
                ) {
                    return Err(invalid("drawing anchor cannot be empty"));
                }
                if element.local_name().as_ref() != b"absoluteAnchor" {
                    return Err(unsupported("copying unknown drawing root children"));
                }
            },
            Event::End(element) => {
                if depth == 2
                    && xdr(&namespace)
                    && matches!(
                        element.local_name().as_ref(),
                        b"twoCellAnchor" | b"oneCellAnchor"
                    )
                {
                    let anchor_start = open_anchor
                        .take()
                        .ok_or_else(|| invalid("drawing anchor close has no start"))?;
                    if anchors.len() >= MAX_ANCHORS {
                        return Err(invalid("worksheet drawing anchor limit exceeded"));
                    }
                    anchors.push(AnchorSpan {
                        start: anchor_start,
                        end,
                    });
                }
                if depth == 1 {
                    root_close_start = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected drawing XML end element"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "drawing transfer rejects DTD and processing instructions",
                ));
            },
            Event::Eof => break,
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0 || open_anchor.is_some() {
        return Err(invalid("incomplete worksheet drawing XML"));
    }
    Ok(DrawingLayout {
        root_open_end: root_open_end.ok_or_else(|| invalid("drawing XML has no root"))?,
        root_close_start: root_close_start
            .ok_or_else(|| invalid("drawing XML has no root close"))?,
        anchors,
    })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Marker {
    None,
    From,
    To,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Coordinate {
    Column,
    Row,
}

fn parse_anchor(xml: &[u8], span: &AnchorSpan) -> Result<Anchor> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut marker = Marker::None;
    let mut coordinate = None;
    let mut from_row = None;
    let mut from_column = None;
    let mut relationship_ids = HashSet::new();
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if end <= span.start {
            continue;
        }
        if start >= span.end {
            break;
        }
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(xml_error)?;
                    if relationship_namespace(&resolver.resolve_attribute(attribute.key).0) {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                            .map_err(xml_error)?
                            .into_owned();
                        relationship_ids.insert(value);
                    }
                }
                if xdr(&namespace) {
                    match element.local_name().as_ref() {
                        b"from" => marker = Marker::From,
                        b"to" => marker = Marker::To,
                        b"col" if marker != Marker::None => coordinate = Some(Coordinate::Column),
                        b"row" if marker != Marker::None => coordinate = Some(Coordinate::Row),
                        _ => {},
                    }
                }
            },
            Event::Text(text) => {
                let Some(coordinate) = coordinate else {
                    continue;
                };
                let value = text
                    .decode()
                    .map_err(xml_error)?
                    .trim()
                    .parse::<u32>()
                    .map_err(|_source| {
                        invalid("drawing anchor coordinate is not an unsigned integer")
                    })?;
                if marker == Marker::From {
                    match coordinate {
                        Coordinate::Column => from_column = Some(value),
                        Coordinate::Row => from_row = Some(value),
                    }
                }
            },
            Event::End(element) if xdr(&namespace) => match element.local_name().as_ref() {
                b"from" | b"to" => marker = Marker::None,
                b"col" | b"row" => coordinate = None,
                _ => {},
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "drawing anchor rejects DTD and processing instructions",
                ));
            },
            Event::Eof => break,
            Event::End(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    let mut relationship_ids = relationship_ids.into_iter().collect::<Vec<_>>();
    relationship_ids.sort_unstable();
    Ok(Anchor {
        from_row: from_row.ok_or_else(|| invalid("drawing anchor has no from row"))?,
        from_column: from_column.ok_or_else(|| invalid("drawing anchor has no from column"))?,
        relationship_ids,
    })
}

fn translate_anchor(
    xml: &[u8],
    span: &AnchorSpan,
    row_delta: i64,
    column_delta: i64,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::new());
    let mut marker = Marker::None;
    let mut coordinate = None;
    loop {
        let start = position(&reader)?;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, _) = resolver.resolve_event(event.clone());
        if end <= span.start {
            continue;
        }
        if start >= span.end {
            break;
        }
        match &event {
            Event::Start(element) if xdr(&namespace) => match element.local_name().as_ref() {
                b"from" => marker = Marker::From,
                b"to" => marker = Marker::To,
                b"col" if marker != Marker::None => coordinate = Some(Coordinate::Column),
                b"row" if marker != Marker::None => coordinate = Some(Coordinate::Row),
                _ => {},
            },
            Event::Text(text) if coordinate.is_some() => {
                let value = text
                    .decode()
                    .map_err(xml_error)?
                    .trim()
                    .parse::<i64>()
                    .map_err(|_source| invalid("drawing anchor coordinate is not an integer"))?;
                let translated = value
                    .checked_add(match coordinate {
                        Some(Coordinate::Column) => column_delta,
                        Some(Coordinate::Row) => row_delta,
                        None => 0,
                    })
                    .ok_or_else(|| invalid("drawing anchor translation overflow"))?;
                let limit = match coordinate {
                    Some(Coordinate::Column) => i64::from(litchi_sheet::COLUMNS),
                    Some(Coordinate::Row) => i64::from(litchi_sheet::ROWS),
                    None => 0,
                };
                if !(0..limit).contains(&translated) {
                    return Err(invalid(
                        "translated drawing anchor exceeds the worksheet grid",
                    ));
                }
                let translated = translated.to_string();
                writer
                    .write_event(Event::Text(BytesText::new(&translated)))
                    .map_err(xml_error)?;
                continue;
            },
            Event::End(element) if xdr(&namespace) => match element.local_name().as_ref() {
                b"from" | b"to" => marker = Marker::None,
                b"col" | b"row" => coordinate = None,
                _ => {},
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        writer.write_event(event).map_err(xml_error)?;
    }
    Ok(writer.into_inner())
}

struct WorksheetLayout {
    root_name: String,
    relationship_namespace: &'static [u8],
    insertion: usize,
    drawing_reference: Option<String>,
    alternate_content: bool,
}

#[derive(Default)]
struct WorksheetChildState {
    successor: Option<usize>,
    drawing_reference: Option<String>,
    alternate_content: bool,
}

fn worksheet_drawing_reference(xml: &[u8]) -> Result<Option<String>> {
    Ok(worksheet_layout(xml)?.drawing_reference)
}

fn worksheet_layout(xml: &[u8]) -> Result<WorksheetLayout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_name = None;
    let mut root_relationship_namespace = None;
    let mut root_close = None;
    let mut child_state = WorksheetChildState::default();
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if !sml(&namespace) || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("drawing transfer requires one worksheet root"));
                    }
                    let qualified_name = element.name();
                    root_name = Some(
                        std::str::from_utf8(qualified_name.as_ref())
                            .map_err(|error| {
                                invalid(format!("worksheet name is not UTF-8: {error}"))
                            })?
                            .to_owned(),
                    );
                    root_relationship_namespace = Some(match namespace {
                        ResolveResult::Bound(Namespace(value)) if value == SML => REL,
                        ResolveResult::Bound(Namespace(value)) if value == STRICT_SML => STRICT_REL,
                        ResolveResult::Unbound
                        | ResolveResult::Bound(_)
                        | ResolveResult::Unknown(_) => {
                            return Err(invalid("drawing transfer requires a worksheet root"));
                        },
                    });
                } else if depth == 1 {
                    inspect_worksheet_child(
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        start,
                        &mut child_state,
                    )?;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
            },
            Event::Empty(element) if depth == 1 => inspect_worksheet_child(
                &namespace,
                &element,
                decoder,
                &resolver,
                start,
                &mut child_state,
            )?,
            Event::End(_) => {
                if depth == 1 {
                    root_close = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected worksheet XML end element"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "drawing transfer rejects DTD and processing instructions",
                ));
            },
            Event::Eof => break,
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(WorksheetLayout {
        root_name: root_name.ok_or_else(|| invalid("worksheet XML has no root"))?,
        relationship_namespace: root_relationship_namespace
            .ok_or_else(|| invalid("worksheet XML has no relationship namespace"))?,
        insertion: child_state
            .successor
            .or(root_close)
            .ok_or_else(|| invalid("worksheet XML has no drawing insertion point"))?,
        drawing_reference: child_state.drawing_reference,
        alternate_content: child_state.alternate_content,
    })
}

fn inspect_worksheet_child(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
    start: usize,
    state: &mut WorksheetChildState,
) -> Result<()> {
    if matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
        && element.local_name().as_ref() == b"AlternateContent"
    {
        state.alternate_content = true;
    }
    if !sml(namespace) {
        return Ok(());
    }
    let local = element.local_name();
    if local.as_ref() == b"drawing" {
        if state.drawing_reference.is_some() {
            return Err(invalid("worksheet has duplicate drawing references"));
        }
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(xml_error)?;
            if relationship_namespace(&resolver.resolve_attribute(attribute.key).0)
                && attribute.key.local_name().as_ref() == b"id"
            {
                state.drawing_reference = Some(
                    attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                        .map_err(xml_error)?
                        .into_owned(),
                );
            }
        }
        if state.drawing_reference.is_none() {
            return Err(invalid(
                "worksheet drawing reference has no relationship ID",
            ));
        }
    } else if drawing_successor(local.as_ref()) && state.successor.is_none() {
        state.successor = Some(start);
    }
    Ok(())
}

fn drawing_successor(local: &[u8]) -> bool {
    matches!(
        local,
        b"legacyDrawing"
            | b"legacyDrawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    )
}

fn allocate_relationship_id(part: &dyn Part) -> Result<String> {
    for index in 1..=MAX_URI_ATTEMPTS {
        let candidate = format!("rIdDrawingCopy{index}");
        if part.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("cannot allocate worksheet drawing relationship ID"))
}

fn allocate_uri(original: &PackURI, reserved: &mut BTreeSet<String>) -> Result<PackURI> {
    let value = original.as_str();
    let (stem, extension) = value
        .rsplit_once('.')
        .map_or((value, ""), |(stem, extension)| (stem, extension));
    for index in 1..=MAX_URI_ATTEMPTS {
        let candidate = if extension.is_empty() {
            format!("{stem}_copy{index}")
        } else {
            format!("{stem}_copy{index}.{extension}")
        };
        if reserved.insert(candidate.clone()) {
            return PackURI::new(&candidate).map_err(invalid);
        }
    }
    Err(invalid(format!(
        "cannot allocate cloned part name for '{original}'"
    )))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source| invalid("XML position does not fit usize"))
}

fn xdr(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == XDR || *value == STRICT_XDR)
}

fn sml(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == SML || *value == STRICT_SML)
}

fn relationship_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == REL || *value == STRICT_REL)
}

fn unsupported(feature: &'static str) -> Error {
    Error::Unsupported { feature }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!("drawing transfer XML error: {error}"))
}
