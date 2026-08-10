//! Lossless, collision-safe transfer of ordinary worksheet drawing graphs.

#![deny(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::float_cmp,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unnecessary_unwrap,
    clippy::wildcard_enum_match_arm,
    reason = "drawing transfer uses checked identifiers, bounded graphs, and explicit ambiguity refusals"
)]

use crate::Workbook;
use crate::package::error::{Error, Result};
use crate::raw::{Header, Limits as RawLimits, Records, Writer as BinaryWriter, kind};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part, TargetMode};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;
use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::fmt;

const XDR: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_ANCHORS: usize = 100_000;
const MAX_GRAPH_PARTS: usize = 4_096;
const MAX_GRAPH_BYTES: usize = 256 * 1024 * 1024;
const MAX_URI_ATTEMPTS: u32 = 10_000;

/// A semantic drawing graph that cannot be copied without guessing ownership.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum DrawingTransferRefusal {
    /// The selected worksheet has no standard drawing.
    SourceDrawingMissing,
    /// The zero-based top-level anchor selector does not exist.
    AnchorMissing(usize),
    /// Pictures use the decoded image-transfer API because their anchor is target-owned.
    PictureUsesImageTransfer,
    /// A graphic frame does not contain an ordinary DrawingML chart.
    NonChartGraphicFrame,
    /// An OLE or otherwise unmodeled object occurs in the selected graph.
    ForeignObject,
    /// Markup-compatibility branches make raw-anchor ownership ambiguous.
    MarkupCompatibility,
    /// Strict and Transitional DrawingML namespaces cannot be mixed losslessly.
    ConformanceMismatch,
    /// A shape/group graph carries package relationships outside ordinary shape ownership.
    RelationshipBearingShape,
    /// A non-visual object has no usable numeric identity.
    MissingObjectId,
    /// A non-visual or connector identity is not an unsigned integer.
    InvalidObjectId(String),
    /// More than one object claims a connector identity.
    AmbiguousObjectId(u32),
    /// A connector endpoint does not resolve to a transferable shape/group anchor.
    UnresolvedConnectorEndpoint(u32),
    /// The chart relationship is missing, external, or has the wrong type.
    InvalidChartRelationship,
    /// A chart-owned internal relationship escapes the bounded inert resource graph.
    WorkbookGlobalChartDependency(String),
    /// A chart relationship owns or links an active embedded object.
    ActiveChartDependency(String),
    /// The source chart graph exceeds the finite part or byte policy.
    ChartGraphLimit,
}

impl fmt::Display for DrawingTransferRefusal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::SourceDrawingMissing => formatter.write_str("source worksheet has no drawing"),
            Self::AnchorMissing(index) => write!(formatter, "source drawing has no anchor {index}"),
            Self::PictureUsesImageTransfer => {
                formatter.write_str("picture anchors must use transfer_image")
            },
            Self::NonChartGraphicFrame => {
                formatter.write_str("graphic frame is not an ordinary chart")
            },
            Self::ForeignObject => formatter.write_str("drawing graph contains a foreign object"),
            Self::MarkupCompatibility => {
                formatter.write_str("drawing graph contains markup-compatibility branches")
            },
            Self::ConformanceMismatch => {
                formatter.write_str("source and target drawing conformance differ")
            },
            Self::RelationshipBearingShape => {
                formatter.write_str("shape/group graph carries package relationships")
            },
            Self::MissingObjectId => {
                formatter.write_str("drawing graph contains a missing object ID")
            },
            Self::InvalidObjectId(value) => {
                write!(formatter, "drawing object ID is invalid: {value}")
            },
            Self::AmbiguousObjectId(id) => {
                write!(formatter, "drawing object ID {id} has ambiguous ownership")
            },
            Self::UnresolvedConnectorEndpoint(id) => {
                write!(
                    formatter,
                    "connector endpoint {id} has no transferable owner"
                )
            },
            Self::InvalidChartRelationship => {
                formatter.write_str("chart frame has no unique internal chart relationship")
            },
            Self::WorkbookGlobalChartDependency(part) => {
                write!(formatter, "chart relationship escapes to {part}")
            },
            Self::ActiveChartDependency(relationship_type) => {
                write!(
                    formatter,
                    "chart relationship is active content: {relationship_type}"
                )
            },
            Self::ChartGraphLimit => formatter.write_str("chart resource graph exceeds limits"),
        }
    }
}

impl std::error::Error for DrawingTransferRefusal {}

#[derive(Debug)]
struct LocatedDrawing {
    uri: PackURI,
}

#[derive(Debug)]
struct AnchorInfo {
    start: usize,
    end: usize,
    ids: Vec<(u32, Option<String>)>,
    endpoints: Vec<u32>,
    relationship_references: Vec<RelationshipReference>,
    foreign_descendant: bool,
}

#[derive(Debug)]
struct RelationshipReference {
    value: String,
    chart_reference: bool,
}

#[derive(Debug)]
struct DrawingLayout {
    strict: bool,
    namespaces: BTreeMap<String, String>,
    anchors: Vec<AnchorInfo>,
}

#[derive(Debug)]
struct ChartGraph {
    root: PackURI,
    parts: Vec<PackURI>,
}

#[derive(Debug)]
struct SourcePlan {
    drawing_uri: PackURI,
    xml: Vec<u8>,
    layout: DrawingLayout,
    selected: BTreeSet<usize>,
    chart_relationship: Option<String>,
}

#[derive(Debug)]
struct TargetPlan {
    location: Option<LocatedDrawing>,
    uri: PackURI,
    xml: Vec<u8>,
    layout: DrawingLayout,
}

#[derive(Debug)]
struct TargetChart {
    relationship_id: String,
    uri: PackURI,
    relationship_type: String,
}

pub(super) fn transfer(
    target: &mut Workbook,
    source_bytes: &[u8],
    source_sheet: usize,
    source_anchor: usize,
    target_sheet: usize,
) -> Result<()> {
    let source = super::root::validated_workbook(source_bytes)?;
    let source_plan = plan_source(&source, source_sheet, source_anchor)?;
    let target_plan = plan_target(target, target_sheet)?;
    if source_plan.layout.strict != target_plan.layout.strict {
        return Err(refused(DrawingTransferRefusal::ConformanceMismatch));
    }
    let (id_mapping, name_mapping) = collision_mapping(
        &source_plan.layout,
        &source_plan.selected,
        &target_plan.layout,
    )?;
    let mut package = target.package.clone();
    let target_chart = copy_selected_chart(&source, &mut package, &source_plan, &target_plan)?;
    let fragments = rewrite_selected_anchors(
        &source_plan,
        &id_mapping,
        &name_mapping,
        target_chart
            .as_ref()
            .map(|chart| chart.relationship_id.as_str()),
    )?;
    let drawing_xml =
        crate::package::drawing_write::append_drawing_anchors(&target_plan.xml, &fragments)?;
    install_transfer(
        &mut package,
        target,
        target_sheet,
        &target_plan,
        drawing_xml,
        target_chart,
    )?;
    package.unsign();
    *target = Workbook::from_opc_package(package)?;
    validate_readback(
        target,
        target_sheet,
        &target_plan,
        source_plan.selected.len(),
    )
}

fn plan_source(source: &Workbook, source_sheet: usize, source_anchor: usize) -> Result<SourcePlan> {
    let source_location = locate_drawing(source, source_sheet)?
        .ok_or_else(|| refused(DrawingTransferRefusal::SourceDrawingMissing))?;
    let source_part = source.package.get_part(&source_location.uri)?;
    let source_inventory = source.sheet_drawing(source_sheet).ok_or_else(|| {
        Error::InvalidFormat("source BrtDrawing has no decoded drawing inventory".to_string())
    })?;
    let layout = drawing_layout(source_part.blob())?;
    if layout.anchors.len() != source_inventory.drawing.anchors.len() {
        return Err(refused(DrawingTransferRefusal::MarkupCompatibility));
    }
    let selected_inventory = source_inventory
        .drawing
        .anchors
        .get(source_anchor)
        .ok_or_else(|| refused(DrawingTransferRefusal::AnchorMissing(source_anchor)))?;

    let mut selected = BTreeSet::new();
    let mut chart_relationship = None;
    match &selected_inventory.object {
        crate::package::drawing::Object::Picture { .. } => {
            return Err(refused(DrawingTransferRefusal::PictureUsesImageTransfer));
        },
        crate::package::drawing::Object::GraphicFrame(frame) => {
            if !frame.is_chart() {
                return Err(refused(DrawingTransferRefusal::NonChartGraphicFrame));
            }
            let relationship_id = frame
                .rel_id
                .as_ref()
                .ok_or_else(|| refused(DrawingTransferRefusal::InvalidChartRelationship))?;
            validate_chart_anchor(
                layout
                    .anchors
                    .get(source_anchor)
                    .ok_or_else(|| refused(DrawingTransferRefusal::AnchorMissing(source_anchor)))?,
                relationship_id,
            )?;
            selected.insert(source_anchor);
            chart_relationship = Some(relationship_id.clone());
        },
        crate::package::drawing::Object::Shape(_)
        | crate::package::drawing::Object::ConnectionShape(_)
        | crate::package::drawing::Object::GroupShape(_) => {
            selected = shape_graph(source_inventory, &layout, source_anchor)?;
        },
    }
    Ok(SourcePlan {
        drawing_uri: source_location.uri.clone(),
        xml: source_part.blob().to_vec(),
        layout,
        selected,
        chart_relationship,
    })
}

fn plan_target(target: &Workbook, target_sheet: usize) -> Result<TargetPlan> {
    let target_location = locate_drawing(target, target_sheet)?;
    ensure_drawing_ownership(target, target_sheet, target_location.as_ref())?;
    let (target_xml, target_uri) = if let Some(location) = &target_location {
        (
            target.package.get_part(&location.uri)?.blob().to_vec(),
            location.uri.clone(),
        )
    } else {
        let uri = allocate_new_drawing_uri(&target.package)?;
        (empty_drawing_xml(is_strict(&target.package)), uri)
    };
    let target_layout = drawing_layout(&target_xml)?;
    Ok(TargetPlan {
        location: target_location,
        uri: target_uri,
        xml: target_xml,
        layout: target_layout,
    })
}

fn copy_selected_chart(
    source: &Workbook,
    package: &mut litchi_opc::OpcPackage,
    source_plan: &SourcePlan,
    target_plan: &TargetPlan,
) -> Result<Option<TargetChart>> {
    let Some(source_id) = &source_plan.chart_relationship else {
        return Ok(None);
    };
    let source_part = source.package.get_part(&source_plan.drawing_uri)?;
    let source_relationship = source_part
        .rels()
        .get(source_id)
        .ok_or_else(|| refused(DrawingTransferRefusal::InvalidChartRelationship))?;
    if source_relationship.is_external()
        || !matches!(source_relationship.reltype(), rt::CHART | rt::STRICT_CHART)
    {
        return Err(refused(DrawingTransferRefusal::InvalidChartRelationship));
    }
    let graph = collect_chart_graph(source, source_relationship.target_partname()?)?;
    let mapping = copy_chart_graph(source, package, &graph)?;
    let copied_chart_uri = mapping.get(graph.root.as_str()).ok_or_else(|| {
        Error::InvalidFormat("copied chart graph has no root mapping".to_string())
    })?;
    let relationship_id = if target_plan.location.is_some() {
        package.get_part_mut(&target_plan.uri)?.relate_to(
            &copied_chart_uri.relative_ref(target_plan.uri.base_uri()),
            source_relationship.reltype(),
        )
    } else {
        "rId1".to_string()
    };
    Ok(Some(TargetChart {
        relationship_id,
        uri: copied_chart_uri.clone(),
        relationship_type: source_relationship.reltype().to_string(),
    }))
}

fn rewrite_selected_anchors(
    source: &SourcePlan,
    id_mapping: &BTreeMap<u32, u32>,
    name_mapping: &BTreeMap<u32, String>,
    target_chart_relationship: Option<&str>,
) -> Result<Vec<Vec<u8>>> {
    let mut fragments = Vec::new();
    for index in &source.selected {
        let info = source.layout.anchors.get(*index).ok_or_else(|| {
            Error::InvalidFormat("selected drawing anchor disappeared".to_string())
        })?;
        let fragment = source.xml.get(info.start..info.end).ok_or_else(|| {
            Error::InvalidFormat("drawing anchor range is outside its part".to_string())
        })?;
        fragments.push(rewrite_anchor(
            fragment,
            &source.layout.namespaces,
            id_mapping,
            name_mapping,
            source.chart_relationship.as_deref(),
            target_chart_relationship,
        )?);
    }
    Ok(fragments)
}

fn install_transfer(
    package: &mut litchi_opc::OpcPackage,
    target: &Workbook,
    target_sheet: usize,
    plan: &TargetPlan,
    drawing_xml: Vec<u8>,
    chart: Option<TargetChart>,
) -> Result<()> {
    if let Some(location) = &plan.location {
        package.get_part_mut(&location.uri)?.set_blob(drawing_xml);
    } else {
        let mut drawing_part =
            BlobPart::new(plan.uri.clone(), ct::OFC_DRAWING.to_string(), drawing_xml);
        if let Some(chart) = chart {
            drawing_part.rels_mut().try_add_relationship(
                chart.relationship_type,
                chart.uri.relative_ref(plan.uri.base_uri()),
                chart.relationship_id,
                TargetMode::Internal,
            )?;
        }
        package.try_add_part(Box::new(drawing_part))?;
        attach_drawing_to_worksheet(package, target, target_sheet, &plan.uri)?;
    }
    Ok(())
}

fn validate_readback(
    target: &Workbook,
    target_sheet: usize,
    plan: &TargetPlan,
    added: usize,
) -> Result<()> {
    let after = target.sheet_drawing(target_sheet).ok_or_else(|| {
        Error::InvalidFormat("transferred drawing failed semantic readback".to_string())
    })?;
    let expected = plan
        .layout
        .anchors
        .len()
        .checked_add(added)
        .ok_or(Error::CapacityOverflow {
            resource: "transferred drawing anchor count",
        })?;
    if after.drawing.anchors.len() != expected {
        return Err(Error::InvalidFormat(
            "transferred drawing anchor count failed readback".to_string(),
        ));
    }
    Ok(())
}

fn refused(value: DrawingTransferRefusal) -> Error {
    Error::DrawingTransfer(value)
}

fn locate_drawing(workbook: &Workbook, sheet: usize) -> Result<Option<LocatedDrawing>> {
    let worksheet_uri = workbook.worksheet_uri(sheet)?;
    let worksheet = workbook.package.get_part(&worksheet_uri)?;
    let mut relationship_id = None;
    for item in Records::new(worksheet.blob()) {
        let record = item?;
        if record.kind() == kind::DRAWING {
            if relationship_id.is_some() {
                return Err(Error::InvalidFormat(
                    "worksheet contains multiple BrtDrawing records".to_string(),
                ));
            }
            let mut cursor = crate::raw::Cursor::new(record.payload(), "BrtDrawing");
            relationship_id = Some(cursor.read_wide_string()?);
        }
    }
    let Some(relationship_id) = relationship_id else {
        return Ok(None);
    };
    let relationship = worksheet.rels().get(&relationship_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "BrtDrawing relationship {relationship_id:?} is absent"
        ))
    })?;
    if relationship.is_external()
        || !matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING)
    {
        return Err(Error::InvalidFormat(
            "BrtDrawing relationship is external or has the wrong type".to_string(),
        ));
    }
    Ok(Some(LocatedDrawing {
        uri: relationship.target_partname()?,
    }))
}

fn ensure_drawing_ownership(
    workbook: &Workbook,
    sheet: usize,
    location: Option<&LocatedDrawing>,
) -> Result<()> {
    if location.is_some() {
        return Ok(());
    }
    let worksheet_uri = workbook.worksheet_uri(sheet)?;
    if workbook
        .package
        .get_part(&worksheet_uri)?
        .rels()
        .iter()
        .any(|relationship| matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING))
    {
        return Err(Error::InvalidFormat(
            "worksheet has a Drawing relationship without BrtDrawing ownership".to_string(),
        ));
    }
    Ok(())
}

fn drawing_layout(xml: &[u8]) -> Result<DrawingLayout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut layout = LayoutBuilder::default();
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, resolved_event) = resolver.resolve_event(event);
        match &resolved_event {
            Event::Start(element) => {
                layout.start(&namespace, element, decoder, &resolver, start)?;
            },
            Event::Empty(element) => inspect_element(
                &namespace,
                element,
                decoder,
                &resolver,
                layout.depth,
                layout.current.as_mut(),
            )?,
            Event::End(element) => layout.end(&namespace, element, end)?,
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "drawing transfer rejects DTD and processing instructions".to_string(),
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    layout.finish()
}

#[derive(Default)]
struct LayoutBuilder {
    depth: usize,
    root_seen: bool,
    strict: bool,
    namespaces: BTreeMap<String, String>,
    open_anchor: Option<usize>,
    current: Option<AnchorInfo>,
    anchors: Vec<AnchorInfo>,
}

impl LayoutBuilder {
    fn start(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: quick_xml::Decoder,
        resolver: &NamespaceResolver,
        start: usize,
    ) -> Result<()> {
        if self.depth == 0 {
            if !xdr(namespace) || element.local_name().as_ref() != b"wsDr" {
                return Err(Error::InvalidFormat(
                    "drawing transfer requires one xdr:wsDr root".to_string(),
                ));
            }
            self.root_seen = true;
            self.strict =
                matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == STRICT_XDR);
            self.namespaces = namespace_declarations(element, decoder)?;
        } else if self.depth == 1 && xdr(namespace) && anchor_element(element) {
            if self.anchors.len() >= MAX_ANCHORS {
                return Err(Error::InvalidLength {
                    expected: MAX_ANCHORS,
                    found: self.anchors.len().saturating_add(1),
                });
            }
            self.open_anchor = Some(start);
            self.current = Some(AnchorInfo {
                start,
                end: 0,
                ids: Vec::new(),
                endpoints: Vec::new(),
                relationship_references: Vec::new(),
                foreign_descendant: false,
            });
        }
        inspect_element(
            namespace,
            element,
            decoder,
            resolver,
            self.depth,
            self.current.as_mut(),
        )?;
        self.depth = self
            .depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("drawing XML nesting overflow".to_string()))?;
        Ok(())
    }

    fn end(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &quick_xml::events::BytesEnd<'_>,
        end: usize,
    ) -> Result<()> {
        if self.depth == 2 && xdr(namespace) && anchor_end(element) {
            let anchor_start = self.open_anchor.take().ok_or_else(|| {
                Error::InvalidFormat("drawing anchor close has no start".to_string())
            })?;
            let mut info = self.current.take().ok_or_else(|| {
                Error::InvalidFormat("drawing anchor has no analysis state".to_string())
            })?;
            if info.start != anchor_start {
                return Err(Error::InvalidFormat(
                    "drawing anchor range ownership changed".to_string(),
                ));
            }
            info.end = end;
            self.anchors.push(info);
        }
        self.depth = self
            .depth
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("drawing XML has an unmatched end".to_string()))?;
        Ok(())
    }

    fn finish(self) -> Result<DrawingLayout> {
        if !self.root_seen || self.depth != 0 || self.open_anchor.is_some() {
            return Err(Error::InvalidFormat(
                "drawing XML is incomplete".to_string(),
            ));
        }
        Ok(DrawingLayout {
            strict: self.strict,
            namespaces: self.namespaces,
            anchors: self.anchors,
        })
    }
}

fn anchor_element(element: &BytesStart<'_>) -> bool {
    matches!(
        element.local_name().as_ref(),
        b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
    )
}

fn anchor_end(element: &quick_xml::events::BytesEnd<'_>) -> bool {
    matches!(
        element.local_name().as_ref(),
        b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
    )
}

fn inspect_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    current: Option<&mut AnchorInfo>,
) -> Result<()> {
    if mce(namespace) {
        return Err(refused(DrawingTransferRefusal::MarkupCompatibility));
    }
    let Some(current) = current else {
        return Ok(());
    };
    let local = element.local_name();
    let local = local.as_ref();
    if xdr(namespace) && matches!(local, b"pic" | b"graphicFrame") && depth > 2 {
        current.foreign_descendant = true;
    }
    if xdr(namespace) && local == b"cNvPr" {
        let id = parse_object_id(
            &unqualified_attribute(element, b"id", decoder)?
                .ok_or_else(|| refused(DrawingTransferRefusal::MissingObjectId))?,
        )?;
        if id == 0 {
            return Err(refused(DrawingTransferRefusal::MissingObjectId));
        }
        let name = unqualified_attribute(element, b"name", decoder)?;
        current.ids.push((id, name));
    }
    if local == b"stCxn" || local == b"endCxn" {
        let id = parse_object_id(
            &unqualified_attribute(element, b"id", decoder)?
                .ok_or_else(|| refused(DrawingTransferRefusal::MissingObjectId))?,
        )?;
        current.endpoints.push(id);
    }
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if relationship_namespace(&resolver.resolve_attribute(attribute.key).0) {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            current.relationship_references.push(RelationshipReference {
                value,
                chart_reference: local == b"chart" && attribute.key.local_name().as_ref() == b"id",
            });
        }
    }
    Ok(())
}

fn namespace_declarations(
    element: &BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Result<BTreeMap<String, String>> {
    let mut values = BTreeMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Encoding(format!("namespace name is not UTF-8: {error}")))?;
        if key == "xmlns" || key.starts_with("xmlns:") {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            values.insert(key.to_string(), value);
        }
    }
    Ok(values)
}

fn validate_chart_anchor(info: &AnchorInfo, relationship_id: &str) -> Result<()> {
    if info.foreign_descendant
        || info.relationship_references.len() != 1
        || !info.relationship_references[0].chart_reference
        || info.relationship_references[0].value != relationship_id
    {
        return Err(refused(DrawingTransferRefusal::InvalidChartRelationship));
    }
    Ok(())
}

fn shape_graph(
    source: &crate::package::drawing::SheetDrawing,
    layout: &DrawingLayout,
    source_anchor: usize,
) -> Result<BTreeSet<usize>> {
    let mut owners = BTreeMap::<u32, Option<usize>>::new();
    for (index, info) in layout.anchors.iter().enumerate() {
        for (id, _) in &info.ids {
            match owners.get(id) {
                None => {
                    owners.insert(*id, Some(index));
                },
                Some(_) => {
                    owners.insert(*id, None);
                },
            }
        }
    }
    let mut selected = BTreeSet::new();
    let mut pending = VecDeque::from([source_anchor]);
    while let Some(index) = pending.pop_front() {
        if !selected.insert(index) {
            continue;
        }
        let info = layout
            .anchors
            .get(index)
            .ok_or_else(|| refused(DrawingTransferRefusal::AnchorMissing(index)))?;
        if info.ids.is_empty() {
            return Err(refused(DrawingTransferRefusal::MissingObjectId));
        }
        if info.foreign_descendant || !info.relationship_references.is_empty() {
            return Err(refused(DrawingTransferRefusal::RelationshipBearingShape));
        }
        validate_ordinary_object(source, info)?;
        for endpoint in &info.endpoints {
            match owners.get(endpoint) {
                Some(Some(owner)) => pending.push_back(*owner),
                Some(None) => {
                    return Err(refused(DrawingTransferRefusal::AmbiguousObjectId(
                        *endpoint,
                    )));
                },
                None => {
                    return Err(refused(
                        DrawingTransferRefusal::UnresolvedConnectorEndpoint(*endpoint),
                    ));
                },
            }
        }
    }
    Ok(selected)
}

fn validate_ordinary_object(
    drawing: &crate::package::drawing::SheetDrawing,
    info: &AnchorInfo,
) -> Result<()> {
    let top_id = info
        .ids
        .first()
        .map(|(id, _)| *id)
        .ok_or_else(|| refused(DrawingTransferRefusal::MissingObjectId))?;
    let mut matches = drawing
        .shapes
        .iter()
        .filter(|anchored| object_id(&anchored.object) == Some(top_id));
    let object = matches
        .next()
        .ok_or_else(|| refused(DrawingTransferRefusal::ForeignObject))?;
    if matches.next().is_some() {
        return Err(refused(DrawingTransferRefusal::AmbiguousObjectId(top_id)));
    }
    validate_object_tree(&object.object)
}

fn object_id(object: &crate::shapes::Object) -> Option<u32> {
    match object {
        crate::shapes::Object::Shape(shape) => shape.non_visual.id,
        crate::shapes::Object::ConnectionShape(connection) => connection.non_visual.id,
        crate::shapes::Object::Group(group) => group.non_visual.id,
        crate::shapes::Object::OleObject(object) => object.non_visual.id,
        crate::shapes::Object::Unknown(_) => None,
    }
}

fn validate_object_tree(object: &crate::shapes::Object) -> Result<()> {
    match object {
        crate::shapes::Object::Shape(_) | crate::shapes::Object::ConnectionShape(_) => Ok(()),
        crate::shapes::Object::Group(group) => {
            for child in &group.children {
                validate_object_tree(child)?;
            }
            Ok(())
        },
        crate::shapes::Object::OleObject(_) | crate::shapes::Object::Unknown(_) => {
            Err(refused(DrawingTransferRefusal::ForeignObject))
        },
    }
}

fn collision_mapping(
    source: &DrawingLayout,
    selected: &BTreeSet<usize>,
    target: &DrawingLayout,
) -> Result<(BTreeMap<u32, u32>, BTreeMap<u32, String>)> {
    let mut used_ids = BTreeSet::new();
    let mut used_names = BTreeSet::new();
    for info in &target.anchors {
        for (id, name) in &info.ids {
            used_ids.insert(*id);
            if let Some(name) = name {
                used_names.insert(name.to_lowercase());
            }
        }
    }
    let mut mapping = BTreeMap::new();
    let mut names = BTreeMap::new();
    let mut next_id = 1u32;
    for index in selected {
        let info = source.anchors.get(*index).ok_or_else(|| {
            Error::InvalidFormat("selected drawing anchor disappeared".to_string())
        })?;
        for (source_id, source_name) in &info.ids {
            if mapping.contains_key(source_id) {
                return Err(refused(DrawingTransferRefusal::AmbiguousObjectId(
                    *source_id,
                )));
            }
            while used_ids.contains(&next_id) {
                next_id = next_id.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("drawing object ID space is exhausted".to_string())
                })?;
            }
            mapping.insert(*source_id, next_id);
            used_ids.insert(next_id);
            if let Some(source_name) = source_name {
                let unique = unique_name(source_name, &mut used_names)?;
                names.insert(*source_id, unique);
            }
            next_id = next_id.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("drawing object ID space is exhausted".to_string())
            })?;
        }
    }
    Ok((mapping, names))
}

fn unique_name(source: &str, used: &mut BTreeSet<String>) -> Result<String> {
    if used.insert(source.to_lowercase()) {
        return Ok(source.to_string());
    }
    for suffix in 2..=MAX_URI_ATTEMPTS {
        let candidate = format!("{source} (Imported {suffix})");
        if used.insert(candidate.to_lowercase()) {
            return Ok(candidate);
        }
    }
    Err(Error::InvalidFormat(
        "drawing object name space is exhausted".to_string(),
    ))
}

fn rewrite_anchor(
    xml: &[u8],
    inherited_namespaces: &BTreeMap<String, String>,
    id_mapping: &BTreeMap<u32, u32>,
    name_mapping: &BTreeMap<u32, String>,
    source_chart_relationship: Option<&str>,
    target_chart_relationship: Option<&str>,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut first_element = true;
    loop {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        match event {
            Event::Start(element) => {
                let rewritten = rewrite_element(
                    &element,
                    reader.decoder(),
                    id_mapping,
                    name_mapping,
                    source_chart_relationship,
                    target_chart_relationship,
                    first_element.then_some(inherited_namespaces),
                )?;
                writer
                    .write_event(Event::Start(rewritten))
                    .map_err(xml_error)?;
                first_element = false;
            },
            Event::Empty(element) => {
                let rewritten = rewrite_element(
                    &element,
                    reader.decoder(),
                    id_mapping,
                    name_mapping,
                    source_chart_relationship,
                    target_chart_relationship,
                    first_element.then_some(inherited_namespaces),
                )?;
                writer
                    .write_event(Event::Empty(rewritten))
                    .map_err(xml_error)?;
                first_element = false;
            },
            Event::Eof => break,
            Event::DocType(_) | Event::Decl(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "drawing anchor contains a declaration, DTD, or processing instruction"
                        .to_string(),
                ));
            },
            Event::End(element) => writer.write_event(Event::End(element)).map_err(xml_error)?,
            Event::Text(value) => writer.write_event(Event::Text(value)).map_err(xml_error)?,
            Event::CData(value) => writer.write_event(Event::CData(value)).map_err(xml_error)?,
            Event::Comment(value) => writer
                .write_event(Event::Comment(value))
                .map_err(xml_error)?,
            Event::GeneralRef(value) => writer
                .write_event(Event::GeneralRef(value))
                .map_err(xml_error)?,
        }
    }
    Ok(writer.into_inner())
}

fn rewrite_element(
    source: &BytesStart<'_>,
    decoder: quick_xml::Decoder,
    id_mapping: &BTreeMap<u32, u32>,
    name_mapping: &BTreeMap<u32, String>,
    source_chart_relationship: Option<&str>,
    target_chart_relationship: Option<&str>,
    inherited_namespaces: Option<&BTreeMap<String, String>>,
) -> Result<BytesStart<'static>> {
    let element_name = std::str::from_utf8(source.name().as_ref())
        .map_err(|error| Error::Encoding(format!("drawing element name is not UTF-8: {error}")))?
        .to_string();
    let local = source.local_name();
    let local = local.as_ref();
    let source_object_id = if local == b"cNvPr" {
        match unqualified_attribute(source, b"id", decoder)? {
            Some(value) => Some(parse_object_id(&value)?),
            None => None,
        }
    } else {
        None
    };
    let mut rewritten = BytesStart::new(element_name);
    let mut present = BTreeSet::new();
    for attribute in source.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Encoding(format!("drawing attribute is not UTF-8: {error}")))?
            .to_string();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        present.insert(key.clone());
        let output = if local == b"cNvPr" && attribute.key.local_name().as_ref() == b"id" {
            let source_id = parse_object_id(&value)?;
            id_mapping
                .get(&source_id)
                .ok_or_else(|| refused(DrawingTransferRefusal::MissingObjectId))?
                .to_string()
        } else if local == b"cNvPr" && attribute.key.local_name().as_ref() == b"name" {
            let source_id =
                source_object_id.ok_or_else(|| refused(DrawingTransferRefusal::MissingObjectId))?;
            name_mapping.get(&source_id).cloned().unwrap_or(value)
        } else if matches!(local, b"stCxn" | b"endCxn")
            && attribute.key.local_name().as_ref() == b"id"
        {
            let source_id = parse_object_id(&value)?;
            id_mapping
                .get(&source_id)
                .ok_or_else(|| {
                    refused(DrawingTransferRefusal::UnresolvedConnectorEndpoint(
                        source_id,
                    ))
                })?
                .to_string()
        } else if local == b"chart"
            && attribute.key.local_name().as_ref() == b"id"
            && source_chart_relationship == Some(value.as_str())
        {
            target_chart_relationship
                .ok_or_else(|| refused(DrawingTransferRefusal::InvalidChartRelationship))?
                .to_string()
        } else {
            value
        };
        rewritten.push_attribute((key.as_str(), output.as_str()));
    }
    if let Some(namespaces) = inherited_namespaces {
        for (key, value) in namespaces {
            if !present.contains(key) {
                rewritten.push_attribute((key.as_str(), value.as_str()));
            }
        }
    }
    Ok(rewritten.into_owned())
}

fn collect_chart_graph(source: &Workbook, root: PackURI) -> Result<ChartGraph> {
    let root_part = source.package.get_part(&root)?;
    if root_part.content_type() != ct::DML_CHART
        || !root.as_str().starts_with("/xl/charts/")
        || !std::path::Path::new(root.as_str())
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
    {
        return Err(refused(DrawingTransferRefusal::InvalidChartRelationship));
    }
    let _validated =
        crate::package::chart_resources::parse_chart_resources(&source.package, root_part)?;
    let mut seen = BTreeSet::new();
    let mut parts = Vec::new();
    let mut pending = VecDeque::from([root.clone()]);
    let mut total = 0usize;
    while let Some(uri) = pending.pop_front() {
        if !seen.insert(uri.as_str().to_string()) {
            continue;
        }
        if seen.len() > MAX_GRAPH_PARTS {
            return Err(refused(DrawingTransferRefusal::ChartGraphLimit));
        }
        let part = source.package.get_part(&uri)?;
        parts.push(uri.clone());
        total = total
            .checked_add(part.blob().len())
            .ok_or_else(|| refused(DrawingTransferRefusal::ChartGraphLimit))?;
        if total > MAX_GRAPH_BYTES {
            return Err(refused(DrawingTransferRefusal::ChartGraphLimit));
        }
        let mut relationships = part.rels().iter().collect::<Vec<_>>();
        relationships.sort_unstable_by_key(|relationship| relationship.r_id());
        for relationship in relationships {
            if matches!(
                relationship.reltype(),
                rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT
            ) {
                return Err(refused(DrawingTransferRefusal::ActiveChartDependency(
                    relationship.reltype().to_string(),
                )));
            }
            if relationship.is_external() {
                continue;
            }
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return Err(refused(
                    DrawingTransferRefusal::WorkbookGlobalChartDependency(
                        relationship.target_ref().to_string(),
                    ),
                ));
            }
            let target = relationship.target_partname()?;
            if !chart_owned_uri(&target) {
                return Err(refused(
                    DrawingTransferRefusal::WorkbookGlobalChartDependency(
                        target.as_str().to_string(),
                    ),
                ));
            }
            pending.push_back(target);
        }
    }
    parts.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    Ok(ChartGraph { root, parts })
}

fn chart_owned_uri(uri: &PackURI) -> bool {
    [
        "/xl/charts/",
        "/xl/drawings/",
        "/xl/media/",
        "/xl/embeddings/",
        "/xl/chartResources/",
        "/xl/theme/",
    ]
    .iter()
    .any(|prefix| uri.as_str().starts_with(prefix))
}

fn copy_chart_graph(
    source: &Workbook,
    target: &mut litchi_opc::OpcPackage,
    graph: &ChartGraph,
) -> Result<BTreeMap<String, PackURI>> {
    let mut reserved = target
        .iter_parts()
        .map(|part| part.partname().as_str().to_string())
        .collect::<BTreeSet<_>>();
    let mut mapping = BTreeMap::new();
    for source_uri in &graph.parts {
        mapping.insert(
            source_uri.as_str().to_string(),
            allocate_graph_uri(source_uri, &mut reserved, target)?,
        );
    }
    for source_uri in &graph.parts {
        let source_part = source.package.get_part(source_uri)?;
        let target_uri = mapping.get(source_uri.as_str()).ok_or_else(|| {
            Error::InvalidFormat("chart graph part has no target mapping".to_string())
        })?;
        let mut target_part = BlobPart::new_shared(
            target_uri.clone(),
            source_part.content_type().to_string(),
            source_part.blob_arc(),
        );
        let mut relationships = source_part.rels().iter().collect::<Vec<_>>();
        relationships.sort_unstable_by_key(|relationship| relationship.r_id());
        for relationship in relationships {
            let (target_ref, mode) = if relationship.is_external() {
                (relationship.target_ref().to_string(), TargetMode::External)
            } else {
                let source_target = relationship.target_partname()?;
                let target_target = mapping.get(source_target.as_str()).ok_or_else(|| {
                    refused(DrawingTransferRefusal::WorkbookGlobalChartDependency(
                        source_target.as_str().to_string(),
                    ))
                })?;
                (
                    target_target.relative_ref(target_uri.base_uri()),
                    TargetMode::Internal,
                )
            };
            target_part.rels_mut().try_add_relationship(
                relationship.reltype().to_string(),
                target_ref,
                relationship.r_id().to_string(),
                mode,
            )?;
        }
        target.try_add_part(Box::new(target_part))?;
    }
    Ok(mapping)
}

fn allocate_graph_uri(
    source: &PackURI,
    reserved: &mut BTreeSet<String>,
    package: &litchi_opc::OpcPackage,
) -> Result<PackURI> {
    let value = source.as_str();
    let (stem, extension) = value
        .rsplit_once('.')
        .map_or((value, ""), |(stem, ext)| (stem, ext));
    for suffix in 1..=MAX_URI_ATTEMPTS {
        let candidate = if extension.is_empty() {
            format!("{stem}_import{suffix}")
        } else {
            format!("{stem}_import{suffix}.{extension}")
        };
        let uri = PackURI::new(candidate.clone())?;
        if !reserved.contains(&candidate) && package.validate_new_part_name(&uri).is_ok() {
            reserved.insert(candidate);
            return Ok(uri);
        }
    }
    Err(Error::InvalidFormat(format!(
        "no collision-free part name remains for {value}"
    )))
}

fn allocate_new_drawing_uri(package: &litchi_opc::OpcPackage) -> Result<PackURI> {
    for index in 1..=MAX_URI_ATTEMPTS {
        let uri = PackURI::new(format!("/xl/drawings/drawing{index}.xml"))?;
        if package.validate_new_part_name(&uri).is_ok() {
            return Ok(uri);
        }
    }
    Err(Error::InvalidFormat(
        "no worksheet drawing part name remains".to_string(),
    ))
}

fn empty_drawing_xml(strict: bool) -> Vec<u8> {
    let xdr = if strict {
        "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
    } else {
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    };
    let a = if strict {
        "http://purl.oclc.org/ooxml/drawingml/main"
    } else {
        "http://schemas.openxmlformats.org/drawingml/2006/main"
    };
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"{xdr}\" xmlns:a=\"{a}\"></xdr:wsDr>"
    )
    .into_bytes()
}

fn attach_drawing_to_worksheet(
    package: &mut litchi_opc::OpcPackage,
    workbook: &Workbook,
    sheet: usize,
    drawing_uri: &PackURI,
) -> Result<()> {
    let worksheet_uri = workbook.worksheet_uri(sheet)?;
    let worksheet_source = package.get_part(&worksheet_uri)?.blob().to_vec();
    let strict = is_strict(package);
    let relationship_id = package.get_part_mut(&worksheet_uri)?.relate_to(
        &drawing_uri.relative_ref(worksheet_uri.base_uri()),
        if strict {
            rt::STRICT_DRAWING
        } else {
            rt::DRAWING
        },
    );
    let mut payload = Vec::new();
    BinaryWriter::new(&mut payload).write_wide_string(&relationship_id)?;
    let mut output = Vec::new();
    let mut inserted = false;
    for item in Records::new(&worksheet_source) {
        let record = item?;
        if !inserted && matches!(record.kind(), kind::BEGIN_LIST_PARTS | kind::END_SHEET) {
            BinaryWriter::new(&mut output).write_record(kind::DRAWING, &payload)?;
            inserted = true;
        }
        copy_record(&worksheet_source, &record, &mut output)?;
    }
    if !inserted {
        return Err(Error::InvalidFormat(
            "worksheet has no safe BrtDrawing insertion boundary".to_string(),
        ));
    }
    package.get_part_mut(&worksheet_uri)?.set_blob(output);
    Ok(())
}

fn copy_record(source: &[u8], record: &crate::raw::Record<'_>, output: &mut Vec<u8>) -> Result<()> {
    let record_source = source.get(record.offset()..).ok_or_else(|| {
        Error::InvalidFormat("record offset is outside worksheet.bin".to_string())
    })?;
    let (_, header_len) = Header::parse(record_source, RawLimits::DEFAULT)?;
    let end = record
        .offset()
        .checked_add(header_len)
        .and_then(|offset| offset.checked_add(record.len()))
        .ok_or(Error::CapacityOverflow {
            resource: "worksheet record range",
        })?;
    output.extend_from_slice(source.get(record.offset()..end).ok_or_else(|| {
        Error::InvalidFormat("record range is outside worksheet.bin".to_string())
    })?);
    Ok(())
}

fn unqualified_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::Decoder,
) -> Result<Option<String>> {
    litchi_ooxml_common::xml::unqualified_attribute_value(element, name, decoder)
        .map_err(Error::from)
}

fn parse_object_id(value: &str) -> Result<u32> {
    value
        .parse::<u32>()
        .map_err(|error| refused(DrawingTransferRefusal::InvalidObjectId(error.to_string())))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|error| {
        Error::InvalidFormat(format!("drawing XML position exceeds usize: {error}"))
    })
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Encoding(format!("invalid drawing XML: {error}"))
}

fn xdr(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == XDR || *value == STRICT_XDR)
}

fn relationship_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == REL || *value == STRICT_REL)
}

fn mce(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn is_strict(package: &litchi_opc::OpcPackage) -> bool {
    package
        .iter_parts()
        .flat_map(|part| part.rels().iter())
        .any(|relationship| {
            relationship
                .reltype()
                .starts_with("http://purl.oclc.org/ooxml/")
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn anchor(ids: &[u32]) -> AnchorInfo {
        AnchorInfo {
            start: 0,
            end: 0,
            ids: ids
                .iter()
                .map(|id| (*id, Some(format!("Object {id}"))))
                .collect(),
            endpoints: Vec::new(),
            relationship_references: Vec::new(),
            foreign_descendant: false,
        }
    }

    #[test]
    fn selected_duplicate_non_visual_identity_is_a_typed_refusal() {
        let source = DrawingLayout {
            strict: false,
            namespaces: BTreeMap::new(),
            anchors: vec![anchor(&[7, 7])],
        };
        let target = DrawingLayout {
            strict: false,
            namespaces: BTreeMap::new(),
            anchors: Vec::new(),
        };
        let result = collision_mapping(&source, &BTreeSet::from([0]), &target);
        assert!(matches!(
            result,
            Err(Error::DrawingTransfer(
                DrawingTransferRefusal::AmbiguousObjectId(7)
            ))
        ));
    }
}
