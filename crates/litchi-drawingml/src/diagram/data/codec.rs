//! `DrawingML` diagram data-model XML parsing and canonical writing.

use crate::{Error, Result};
use litchi_ooxml_common::mce::process_str;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt;

use super::validation::{invalid, limit, xml_error};
use super::{
    Conformance, Connection, ConnectionType, DGM_NAMESPACE, DGM_NAMESPACE_STRICT, DML_NAMESPACE,
    DML_NAMESPACE_STRICT, DiagramDataModel, Id, MAX_CONNECTIONS, MAX_DATA_MODEL_XML, MAX_DEPTH,
    MAX_NODES, MAX_POINTS, MAX_TEXT_BYTES, Point, PointType,
};

impl DiagramDataModel {
    /// Parse a `dgm:dataModel` document (transitional or Strict namespace).
    ///
    /// The input is first rewritten by markup-compatibility processing so
    /// `mc:AlternateContent` wrappers resolve to their fallback content.
    /// Unmodeled formatting and extension content is validated as XML structure
    /// but is not retained; see the publication warning on [`Self::to_xml`].
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn parse(xml: &str) -> Result<Self> {
        if xml.len() > MAX_DATA_MODEL_XML {
            return Err(limit("data-model XML bytes"));
        }
        let processed = process_str(xml)?;
        if processed.len() > MAX_DATA_MODEL_XML {
            return Err(limit("processed data-model XML bytes"));
        }
        let mut reader = NsReader::from_reader(processed.as_bytes());
        // Text inside `a:t` is significant: preserve whitespace verbatim and
        // resolve entity references explicitly (`Event::GeneralRef`).
        reader.config_mut().trim_text(false);

        let mut model = DiagramDataModel::default();
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut root_seen = false;
        let mut root_closed = false;
        let mut point_list_seen = false;
        let mut point_list_depth: Option<usize> = None;
        let mut connection_list_seen = false;
        let mut connection_list_depth: Option<usize> = None;
        // Currently open point: depth at which it started plus its builder.
        let mut open_point: Option<(usize, PointBuilder)> = None;
        // Depths of the open `dgm:t` text body and `a:t` text leaf, if any.
        let mut text_body_depth: Option<usize> = None;
        let mut run_text_depth: Option<usize> = None;

        loop {
            match reader.read_event_into(&mut buffer).map_err(xml_error)? {
                Event::Start(element) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("data-model XML depth"))?;
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("data-model XML node count"))?;
                    if nodes > MAX_NODES || depth > MAX_DEPTH {
                        return Err(limit("data-model XML structure"));
                    }
                    let namespace = element_namespace(&reader, &element)?;
                    let local = local_name(&element)?;
                    if depth == 1 {
                        if root_seen || root_closed {
                            return Err(invalid("multiple data-model roots"));
                        }
                        if !is_dgm(&namespace) || local != "dataModel" {
                            return Err(invalid("invalid data-model root or namespace"));
                        }
                        root_seen = true;
                    } else if is_dgm(&namespace) {
                        match local.as_str() {
                            "ptLst" => {
                                if depth != 2 || point_list_seen {
                                    return Err(invalid("invalid or duplicate diagram point list"));
                                }
                                point_list_seen = true;
                                point_list_depth = Some(depth);
                            },
                            "cxnLst" => {
                                if depth != 2
                                    || connection_list_seen
                                    || !point_list_seen
                                    || point_list_depth.is_some()
                                {
                                    return Err(invalid(
                                        "invalid, duplicate, or out-of-order diagram connection list",
                                    ));
                                }
                                connection_list_seen = true;
                                connection_list_depth = Some(depth);
                            },
                            "pt" => {
                                if point_list_depth.is_none_or(|list| depth != list + 1) {
                                    return Err(invalid("diagram point is outside its point list"));
                                }
                                if open_point.is_some() {
                                    return Err(invalid("nested diagram point"));
                                }
                                if model.points.len() >= MAX_POINTS {
                                    return Err(limit("diagram point count"));
                                }
                                open_point = Some((depth, PointBuilder::from_element(&element)?));
                            },
                            "cxn" => {
                                if connection_list_depth.is_none_or(|list| depth != list + 1) {
                                    return Err(invalid(
                                        "diagram connection is outside its connection list",
                                    ));
                                }
                                push_connection(&mut model, &element)?;
                            },
                            "prSet" => {
                                let Some((point_depth, builder)) = &mut open_point else {
                                    return Err(invalid("diagram property set is outside a point"));
                                };
                                if depth != *point_depth + 1 {
                                    return Err(invalid(
                                        "diagram property set is not a direct point child",
                                    ));
                                }
                                builder.read_pr_set(&element)?;
                            },
                            "t" => {
                                let Some((point_depth, builder)) = &mut open_point else {
                                    return Err(invalid("diagram text body is outside a point"));
                                };
                                if depth != *point_depth + 1
                                    || text_body_depth.is_some()
                                    || builder.text_body_seen
                                {
                                    return Err(invalid(
                                        "invalid or duplicate diagram point text body",
                                    ));
                                }
                                builder.text_body_seen = true;
                                text_body_depth = Some(depth);
                            },
                            _ => {},
                        }
                    }
                    if is_dml(&namespace)
                        && local == "t"
                        && text_body_depth.is_some_and(|body| depth > body)
                    {
                        if run_text_depth.is_some() {
                            return Err(invalid("nested DrawingML text leaf"));
                        }
                        run_text_depth = Some(depth);
                    }
                },
                Event::Empty(element) => {
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("data-model XML node count"))?;
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("data-model XML depth"))?;
                    if nodes > MAX_NODES || child_depth > MAX_DEPTH {
                        return Err(limit("data-model XML structure"));
                    }
                    let namespace = element_namespace(&reader, &element)?;
                    let local = local_name(&element)?;
                    if child_depth == 1 {
                        if root_seen || root_closed {
                            return Err(invalid("multiple data-model roots"));
                        }
                        if !is_dgm(&namespace) || local != "dataModel" {
                            return Err(invalid("invalid data-model root or namespace"));
                        }
                        return Err(invalid("diagram data model lacks its required point list"));
                    } else if is_dgm(&namespace) {
                        match local.as_str() {
                            "ptLst" => {
                                if child_depth != 2 || point_list_seen {
                                    return Err(invalid("invalid or duplicate diagram point list"));
                                }
                                point_list_seen = true;
                            },
                            "cxnLst" => {
                                if child_depth != 2
                                    || connection_list_seen
                                    || !point_list_seen
                                    || point_list_depth.is_some()
                                {
                                    return Err(invalid(
                                        "invalid, duplicate, or out-of-order diagram connection list",
                                    ));
                                }
                                connection_list_seen = true;
                            },
                            "pt" => {
                                if point_list_depth.is_none_or(|list| child_depth != list + 1) {
                                    return Err(invalid("diagram point is outside its point list"));
                                }
                                if open_point.is_some() {
                                    return Err(invalid("nested diagram point"));
                                }
                                if model.points.len() >= MAX_POINTS {
                                    return Err(limit("diagram point count"));
                                }
                                let builder = PointBuilder::from_element(&element)?;
                                model.points.push(builder.finish()?);
                            },
                            "cxn" => {
                                if connection_list_depth.is_none_or(|list| child_depth != list + 1)
                                {
                                    return Err(invalid(
                                        "diagram connection is outside its connection list",
                                    ));
                                }
                                push_connection(&mut model, &element)?;
                            },
                            "prSet" => {
                                let Some((point_depth, builder)) = &mut open_point else {
                                    return Err(invalid("diagram property set is outside a point"));
                                };
                                if child_depth != *point_depth + 1 {
                                    return Err(invalid(
                                        "diagram property set is not a direct point child",
                                    ));
                                }
                                builder.read_pr_set(&element)?;
                            },
                            "t" => {
                                let Some((point_depth, builder)) = &mut open_point else {
                                    return Err(invalid("diagram text body is outside a point"));
                                };
                                if child_depth != *point_depth + 1 || builder.text_body_seen {
                                    return Err(invalid(
                                        "invalid or duplicate diagram point text body",
                                    ));
                                }
                                builder.text_body_seen = true;
                            },
                            _ => {},
                        }
                    }
                },
                Event::Text(event) => {
                    if run_text_depth.is_some()
                        && let Some((_, builder)) = &mut open_point
                    {
                        let text = std::str::from_utf8(event.as_ref()).map_err(xml_error)?;
                        let text = quick_xml::escape::unescape(text).map_err(xml_error)?;
                        if builder
                            .text
                            .len()
                            .checked_add(text.len())
                            .is_none_or(|length| length > MAX_TEXT_BYTES)
                        {
                            return Err(limit("diagram text bytes"));
                        }
                        builder.text.push_str(&text);
                    } else if depth == 0
                        && event
                            .as_ref()
                            .iter()
                            .any(|byte| !matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                    {
                        return Err(invalid("text is outside the data-model root"));
                    }
                },
                Event::GeneralRef(reference) => {
                    if run_text_depth.is_some()
                        && let Some((_, builder)) = &mut open_point
                    {
                        let text = litchi_ooxml_common::xml::decode_xml_reference(&reference)?;
                        if builder
                            .text
                            .len()
                            .checked_add(text.len())
                            .is_none_or(|length| length > MAX_TEXT_BYTES)
                        {
                            return Err(limit("diagram text bytes"));
                        }
                        builder.text.push_str(&text);
                    } else if depth == 0 {
                        return Err(invalid("entity reference is outside the data-model root"));
                    }
                },
                Event::End(element) => {
                    let local = local_name_end(&element)?;
                    if run_text_depth == Some(depth) {
                        if local != "t" {
                            return Err(invalid("unbalanced DrawingML text leaf"));
                        }
                        run_text_depth = None;
                    }
                    if text_body_depth == Some(depth) {
                        if local != "t" {
                            return Err(invalid("unbalanced diagram text body"));
                        }
                        if run_text_depth.is_some() {
                            return Err(invalid("unterminated DrawingML text leaf"));
                        }
                        text_body_depth = None;
                    }
                    if open_point
                        .as_ref()
                        .is_some_and(|(start, _)| *start == depth)
                    {
                        if local != "pt" {
                            return Err(invalid("unbalanced diagram point"));
                        }
                        let Some((_, builder)) = open_point.take() else {
                            return Err(invalid("diagram point state is inconsistent"));
                        };
                        model.points.push(builder.finish()?);
                    }
                    if point_list_depth == Some(depth) {
                        if local != "ptLst" {
                            return Err(invalid("unbalanced diagram point list"));
                        }
                        point_list_depth = None;
                    }
                    if connection_list_depth == Some(depth) {
                        if local != "cxnLst" {
                            return Err(invalid("unbalanced diagram connection list"));
                        }
                        connection_list_depth = None;
                    }
                    if depth == 0 {
                        return Err(invalid("unexpected data-model closing element"));
                    }
                    if depth == 1 {
                        if local != "dataModel" || !root_seen {
                            return Err(invalid("unbalanced data-model root"));
                        }
                        if !point_list_seen {
                            return Err(invalid(
                                "diagram data model lacks its required point list",
                            ));
                        }
                        root_closed = true;
                    }
                    depth -= 1;
                },
                Event::DocType(_) => return Err(invalid("DTDs are rejected")),
                Event::CData(_) => return Err(invalid("CDATA is rejected")),
                Event::Eof => break,
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
            }
            buffer.clear();
        }
        if !root_seen
            || !root_closed
            || !point_list_seen
            || depth != 0
            || point_list_depth.is_some()
            || connection_list_depth.is_some()
            || open_point.is_some()
            || text_body_depth.is_some()
            || run_text_depth.is_some()
        {
            return Err(invalid("missing or unterminated data-model root"));
        }
        Ok(model)
    }
    /// Serializes a canonical, validated document from the modeled semantics.
    ///
    /// # Publication safety
    ///
    /// This is not a lossless rewrite API for arbitrary XML passed to
    /// [`Self::parse`]. Rich-text formatting, shape properties, backgrounds,
    /// whole-diagram formatting, and extension lists are outside this model and
    /// will not be emitted. Use this method for fresh authoring or only after
    /// deliberately accepting that canonicalization.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn to_xml(&self, conformance: Conformance) -> Result<String> {
        let capacity = self.serialized_xml_len(conformance)?;
        let mut xml = String::new();
        xml.try_reserve_exact(capacity)
            .map_err(|error| reserve_error(&error))?;
        self.write_validated_xml(&mut xml, conformance)?;
        Ok(xml)
    }

    /// Serializes modeled semantics into a caller-owned sink for allocation
    /// reuse.
    ///
    /// This has the same non-lossless publication contract as [`Self::to_xml`].
    /// Validation and allocation complete before the destination is changed, so
    /// an error leaves its previous contents intact.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn write_xml(&self, xml: &mut String, conformance: Conformance) -> Result<()> {
        let additional = self.serialized_xml_len(conformance)?;
        xml.try_reserve_exact(additional)
            .map_err(|error| reserve_error(&error))?;
        let initial_len = xml.len();
        if let Err(error) = self.write_validated_xml(xml, conformance) {
            xml.truncate(initial_len);
            return Err(error);
        }
        Ok(())
    }
    pub(super) fn write_validated_xml(
        &self,
        xml: &mut impl fmt::Write,
        conformance: Conformance,
    ) -> Result<()> {
        let (diagram_namespace, main_namespace) = match conformance {
            Conformance::Transitional => (DGM_NAMESPACE, DML_NAMESPACE),
            Conformance::Strict => (DGM_NAMESPACE_STRICT, DML_NAMESPACE_STRICT),
        };
        write!(
            xml,
            "<dgm:dataModel xmlns:dgm=\"{diagram_namespace}\" xmlns:a=\"{main_namespace}\"><dgm:ptLst>"
        )
        .map_err(write_error)?;
        for point in &self.points {
            write_point(xml, point)?;
        }
        xml.write_str("</dgm:ptLst><dgm:cxnLst>")
            .map_err(write_error)?;
        for connection in &self.connections {
            write_connection(xml, connection)?;
        }
        xml.write_str("</dgm:cxnLst></dgm:dataModel>")
            .map_err(write_error)
    }
}
fn write_point(xml: &mut impl fmt::Write, point: &Point) -> Result<()> {
    write!(xml, "<dgm:pt modelId=\"{}\"", point.id).map_err(write_error)?;
    match point.kind {
        PointType::Node => {},
        PointType::Document => xml.write_str(" type=\"doc\"").map_err(write_error)?,
        PointType::Assistant => xml.write_str(" type=\"asst\"").map_err(write_error)?,
        PointType::ParentTransition(connection) => {
            write!(xml, " type=\"parTrans\" cxnId=\"{connection}\"").map_err(write_error)?;
        },
        PointType::SiblingTransition(connection) => {
            write!(xml, " type=\"sibTrans\" cxnId=\"{connection}\"").map_err(write_error)?;
        },
        PointType::Presentation => xml.write_str(" type=\"pres\"").map_err(write_error)?,
    }

    let has_property_set = point.layout_type_id.is_some()
        || point.quick_style_type_id.is_some()
        || point.color_style_type_id.is_some();
    if !has_property_set && point.text.is_empty() {
        return xml.write_str("/>").map_err(write_error);
    }
    xml.write_char('>').map_err(write_error)?;
    if has_property_set {
        xml.write_str("<dgm:prSet").map_err(write_error)?;
        write_optional_attribute(xml, "loTypeId", point.layout_type_id.as_deref())?;
        write_optional_attribute(xml, "qsTypeId", point.quick_style_type_id.as_deref())?;
        write_optional_attribute(xml, "csTypeId", point.color_style_type_id.as_deref())?;
        xml.write_str("/>").map_err(write_error)?;
    }
    if !point.text.is_empty() {
        xml.write_str("<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t xml:space=\"preserve\">")
            .map_err(write_error)?;
        write_xml_escaped(xml, &point.text)?;
        xml.write_str("</a:t></a:r></a:p></dgm:t>")
            .map_err(write_error)?;
    }
    xml.write_str("</dgm:pt>").map_err(write_error)
}

fn write_connection(xml: &mut impl fmt::Write, connection: &Connection) -> Result<()> {
    write!(
        xml,
        "<dgm:cxn modelId=\"{}\" srcId=\"{}\" destId=\"{}\" srcOrd=\"{}\" destOrd=\"{}\"",
        connection.id,
        connection.source,
        connection.destination,
        connection.src_ord,
        connection.dest_ord,
    )
    .map_err(write_error)?;
    match &connection.kind {
        ConnectionType::Parent {
            parent_transition,
            sibling_transition,
        } => write!(
            xml,
            " parTransId=\"{parent_transition}\" sibTransId=\"{sibling_transition}\""
        )
        .map_err(write_error)?,
        ConnectionType::Presentation(presentation) => {
            xml.write_str(" type=\"presOf\"").map_err(write_error)?;
            write_optional_attribute(xml, "presId", Some(presentation))?;
        },
        ConnectionType::PresentationParent => {
            xml.write_str(" type=\"presParOf\"").map_err(write_error)?;
        },
        ConnectionType::Unknown => xml
            .write_str(" type=\"unknownRelationship\"")
            .map_err(write_error)?,
    }
    xml.write_str("/>").map_err(write_error)
}

fn write_optional_attribute(
    xml: &mut impl fmt::Write,
    name: &str,
    value: Option<&str>,
) -> Result<()> {
    if let Some(value) = value {
        write!(xml, " {name}=\"").map_err(write_error)?;
        write_xml_escaped(xml, value)?;
        xml.write_char('"').map_err(write_error)?;
    }
    Ok(())
}

fn write_error(error: fmt::Error) -> Error {
    Error::Xml(format!("failed to write diagram XML: {error}"))
}

fn reserve_error(error: &std::collections::TryReserveError) -> Error {
    invalid(format!("cannot reserve diagram XML output: {error}"))
}

fn write_xml_escaped(xml: &mut impl fmt::Write, value: &str) -> Result<()> {
    let mut start = 0usize;
    for (offset, character) in value.char_indices() {
        let replacement = match character {
            '&' => Some("&amp;"),
            '<' => Some("&lt;"),
            '>' => Some("&gt;"),
            '\'' => Some("&apos;"),
            '"' => Some("&quot;"),
            _ => None,
        };
        if let Some(replacement) = replacement {
            xml.write_str(&value[start..offset]).map_err(write_error)?;
            xml.write_str(replacement).map_err(write_error)?;
            start = offset + character.len_utf8();
        }
    }
    xml.write_str(&value[start..]).map_err(write_error)
}
#[derive(Debug, Clone, Copy, Default)]
enum PointTag {
    #[default]
    Node,
    Document,
    Assistant,
    ParentTransition,
    SiblingTransition,
    Presentation,
}

fn parse_point_tag(value: &str) -> Result<PointTag> {
    match xml_token(value) {
        "node" => Ok(PointTag::Node),
        "doc" => Ok(PointTag::Document),
        "asst" => Ok(PointTag::Assistant),
        "parTrans" => Ok(PointTag::ParentTransition),
        "sibTrans" => Ok(PointTag::SiblingTransition),
        "pres" => Ok(PointTag::Presentation),
        value => Err(invalid(format!("invalid diagram point type `{value}`"))),
    }
}

#[derive(Default)]
struct PointBuilder {
    id: Option<Id>,
    tag: PointTag,
    connection: Option<Id>,
    text: String,
    layout_type_id: Option<String>,
    quick_style_type_id: Option<String>,
    color_style_type_id: Option<String>,
    property_set_seen: bool,
    text_body_seen: bool,
}

impl PointBuilder {
    fn from_element(element: &BytesStart<'_>) -> Result<Self> {
        let mut builder = PointBuilder::default();
        for (name, value) in attributes(element)? {
            match name.as_str() {
                "modelId" => builder.id = Some(parse_id(&value, "diagram point modelId")?),
                "type" => builder.tag = parse_point_tag(&value)?,
                "cxnId" => builder.connection = Some(parse_id(&value, "diagram point cxnId")?),
                _ => {},
            }
        }
        if builder.id.is_none() {
            return Err(invalid("diagram point lacks modelId"));
        }
        Ok(builder)
    }

    fn read_pr_set(&mut self, element: &BytesStart<'_>) -> Result<()> {
        if self.property_set_seen {
            return Err(invalid("duplicate diagram point property set"));
        }
        self.property_set_seen = true;
        for (name, value) in attributes(element)? {
            match name.as_str() {
                "loTypeId" => self.layout_type_id = Some(value),
                "qsTypeId" => self.quick_style_type_id = Some(value),
                "csTypeId" => self.color_style_type_id = Some(value),
                _ => {},
            }
        }
        Ok(())
    }

    fn finish(self) -> Result<Point> {
        let id = self
            .id
            .ok_or_else(|| invalid("diagram point lacks modelId"))?;
        let kind = match self.tag {
            PointTag::Node => PointType::Node,
            PointTag::Document => PointType::Document,
            PointTag::Assistant => PointType::Assistant,
            PointTag::ParentTransition => {
                PointType::ParentTransition(self.connection.unwrap_or_else(|| Id::number(0)))
            },
            PointTag::SiblingTransition => {
                PointType::SiblingTransition(self.connection.unwrap_or_else(|| Id::number(0)))
            },
            PointTag::Presentation => PointType::Presentation,
        };
        Ok(Point {
            id,
            kind,
            text: self.text,
            layout_type_id: self.layout_type_id,
            quick_style_type_id: self.quick_style_type_id,
            color_style_type_id: self.color_style_type_id,
        })
    }
}

#[derive(Debug, Clone, Copy, Default)]
enum ConnectionTag {
    #[default]
    Parent,
    Presentation,
    PresentationParent,
    Unknown,
}

fn parse_connection_tag(value: &str) -> Result<ConnectionTag> {
    match xml_token(value) {
        "parOf" => Ok(ConnectionTag::Parent),
        "presOf" => Ok(ConnectionTag::Presentation),
        "presParOf" => Ok(ConnectionTag::PresentationParent),
        "unknownRelationship" => Ok(ConnectionTag::Unknown),
        value => Err(invalid(format!(
            "invalid diagram connection type `{value}`"
        ))),
    }
}

fn push_connection(model: &mut DiagramDataModel, element: &BytesStart<'_>) -> Result<()> {
    if model.connections.len() >= MAX_CONNECTIONS {
        return Err(limit("diagram connection count"));
    }
    let mut id = None;
    let mut tag = ConnectionTag::default();
    let mut source = None;
    let mut destination = None;
    let mut source_order = None;
    let mut destination_order = None;
    let mut parent_transition = None;
    let mut sibling_transition = None;
    let mut presentation = None;
    for (name, value) in attributes(element)? {
        match name.as_str() {
            "modelId" => id = Some(parse_id(&value, "diagram connection modelId")?),
            "type" => tag = parse_connection_tag(&value)?,
            "srcId" => source = Some(parse_id(&value, "diagram connection srcId")?),
            "destId" => destination = Some(parse_id(&value, "diagram connection destId")?),
            "srcOrd" => source_order = Some(parse_order(&value)?),
            "destOrd" => destination_order = Some(parse_order(&value)?),
            "parTransId" => {
                parent_transition = Some(parse_id(&value, "diagram connection parTransId")?);
            },
            "sibTransId" => {
                sibling_transition = Some(parse_id(&value, "diagram connection sibTransId")?);
            },
            "presId" => presentation = Some(value),
            _ => {},
        }
    }
    let id = id.ok_or_else(|| invalid("diagram connection lacks modelId"))?;
    let source = source.ok_or_else(|| invalid("diagram connection lacks srcId"))?;
    let destination = destination.ok_or_else(|| invalid("diagram connection lacks destId"))?;
    let source_order = source_order.ok_or_else(|| invalid("diagram connection lacks srcOrd"))?;
    let destination_order =
        destination_order.ok_or_else(|| invalid("diagram connection lacks destOrd"))?;
    let kind = match tag {
        ConnectionTag::Parent => ConnectionType::Parent {
            parent_transition: parent_transition.unwrap_or_else(|| Id::number(0)),
            sibling_transition: sibling_transition.unwrap_or_else(|| Id::number(0)),
        },
        ConnectionTag::Presentation => {
            ConnectionType::Presentation(presentation.unwrap_or_default())
        },
        ConnectionTag::PresentationParent => ConnectionType::PresentationParent,
        ConnectionTag::Unknown => ConnectionType::Unknown,
    };
    model.connections.push(Connection::new(
        id,
        kind,
        source,
        destination,
        source_order,
        destination_order,
    ));
    Ok(())
}

fn parse_order(value: &str) -> Result<u32> {
    xml_token(value)
        .parse()
        .map_err(|_error| invalid("invalid diagram connection order"))
}

fn parse_id(value: &str, description: &str) -> Result<Id> {
    value
        .parse()
        .map_err(|_error| invalid(format!("invalid {description} `{value}`")))
}

fn xml_token(value: &str) -> &str {
    value.trim_matches([' ', '\t', '\r', '\n'])
}

/// Unnamespaced, unescaped `(local name, value)` attribute pairs.
fn attributes(element: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        // Schema attributes on these elements are unqualified. An extension
        // attribute with the same local name must not impersonate one merely
        // because namespace prefixes are otherwise discarded below.
        if raw.contains(&b':') {
            continue;
        }
        let name = std::str::from_utf8(attribute.key.local_name().as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = std::str::from_utf8(attribute.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(value)
            .map_err(xml_error)?
            .into_owned();
        values.push((name, value));
    }
    Ok(values)
}

fn element_namespace(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
    match reader.resolver().resolve_element(element.name()).0 {
        ResolveResult::Bound(Namespace(namespace)) => Ok(std::str::from_utf8(namespace)
            .map_err(xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn local_name(element: &BytesStart<'_>) -> Result<String> {
    Ok(std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned())
}

fn local_name_end(element: &quick_xml::events::BytesEnd<'_>) -> Result<String> {
    Ok(std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned())
}

fn is_dgm(namespace: &str) -> bool {
    matches!(namespace, DGM_NAMESPACE | DGM_NAMESPACE_STRICT)
}

fn is_dml(namespace: &str) -> bool {
    matches!(namespace, DML_NAMESPACE | DML_NAMESPACE_STRICT)
}

#[cfg(test)]
mod tests {
    use super::write_optional_attribute;

    #[test]
    fn optional_attribute_keeps_minimal_schema_lexical_form() -> crate::Result<()> {
        let mut xml = String::new();
        write_optional_attribute(&mut xml, "presId", None)?;
        assert!(xml.is_empty());

        write_optional_attribute(&mut xml, "presId", Some("a&\"b"))?;
        assert_eq!(xml, " presId=\"a&amp;&quot;b\"");
        Ok(())
    }
}
