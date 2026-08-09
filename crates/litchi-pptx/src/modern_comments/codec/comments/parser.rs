use super::super::super::model::{
    Anchor, AnchorKind, Comment, List, NamespaceDeclaration, Position, Reply, Status,
};
use super::super::super::{
    AC, MAX_BYTES, MAX_COMMENTS, MAX_DEPTH, MAX_NODES, MAX_REPLIES, P188, PC,
};
use super::validation::{
    bounded, invalid, limit, validate_date_time, validate_guid, validate_model, validate_namespaces,
};
use crate::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AnchorFamily {
    Slide,
    Drawing,
    Text,
    Unknown,
}

#[derive(Debug, Clone, Copy)]
enum RawOwner {
    Comment(usize),
    Reply(usize, usize),
}

#[derive(Debug, Clone, Copy)]
enum RawKind {
    Anchor(usize, AnchorKind),
    TextBody(RawOwner),
    Extension(RawOwner),
}

#[derive(Debug)]
enum FrameKind {
    Root,
    Comment {
        index: usize,
        stage: u8,
        anchor_family: Option<AnchorFamily>,
    },
    Position,
    ReplyList {
        comment: usize,
    },
    Reply {
        comment: usize,
        reply: usize,
        stage: u8,
    },
    Raw {
        start: usize,
        kind: RawKind,
    },
    Opaque,
}

#[derive(Debug)]
struct Frame {
    kind: FrameKind,
    namespace: String,
    local: String,
}

pub(super) fn parse_comment_list(xml: &[u8]) -> Result<List> {
    if xml.len() > MAX_BYTES {
        return Err(limit("modern Comment part bytes"));
    }
    let selected = process_ooxml(xml)?;
    if selected.len() > MAX_BYTES {
        return Err(limit("MCE-processed modern Comment bytes"));
    }
    let bytes = selected.as_ref();
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Frame> = Vec::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_prefix = String::new();
    let mut namespace_declarations = Vec::new();
    let mut comments = Vec::new();
    let mut nodes = 0usize;
    let mut reply_count = 0usize;

    loop {
        let start_offset = reader.buffer_position() as usize;
        let decoder = reader.decoder();
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = resolve_namespace(resolved)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("modern Comment nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("modern Comment nodes"));
                }
                if stack.len() + 1 > MAX_DEPTH {
                    return Err(limit("modern Comment XML depth"));
                }
                let local = decode_name(element.local_name().as_ref())?;
                let kind = if stack.is_empty() {
                    if root_seen || root_closed || namespace != P188 || local != "cmLst" {
                        return Err(invalid("modern Comment root must be p188:cmLst"));
                    }
                    root_prefix = element_prefix(&element)?;
                    namespace_declarations =
                        namespace_declarations_from(&element, decoder, Some(&root_prefix))?;
                    no_non_namespace_attributes(&element)?;
                    root_seen = true;
                    FrameKind::Root
                } else {
                    let parent = stack
                        .last_mut()
                        .ok_or_else(|| invalid("modern Comment child is missing its parent"))?;
                    child_frame(
                        &mut comments,
                        &mut reply_count,
                        parent,
                        &element,
                        decoder,
                        &namespace,
                        &local,
                        start_offset,
                    )?
                };
                let frame = Frame {
                    kind,
                    namespace,
                    local,
                };
                if empty {
                    close_empty(&frame.kind)?;
                    attach_raw(
                        &frame.kind,
                        bytes,
                        start_offset,
                        reader.buffer_position() as usize,
                        &mut comments,
                    )?;
                    if matches!(frame.kind, FrameKind::Root) {
                        root_closed = true;
                    }
                } else {
                    stack.push(frame);
                }
            },
            Event::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected modern Comment closing element"))?;
                let local = decode_name(element.local_name().as_ref())?;
                if frame.namespace != namespace || frame.local != local {
                    return Err(invalid("mismatched modern Comment closing element"));
                }
                attach_raw(
                    &frame.kind,
                    bytes,
                    0,
                    reader.buffer_position() as usize,
                    &mut comments,
                )?;
                if matches!(frame.kind, FrameKind::Root) {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                if !inside_raw(&stack) {
                    let decoded = text.decode().map_err(xml_error)?;
                    let value = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                    if !value.trim().is_empty() {
                        return Err(invalid("unexpected text in modern Comment metadata"));
                    }
                }
            },
            Event::CData(text) => {
                if !inside_raw(&stack) && !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("unexpected CDATA in modern Comment metadata"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD, processing instructions, and general references are rejected",
                ));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("unterminated modern Comment part"));
    }
    let value = List {
        root_prefix,
        namespace_declarations,
        comments,
    };
    validate_model(&value)?;
    Ok(value)
}

#[allow(
    clippy::too_many_arguments,
    reason = "parser frame carries one slot per comment child element"
)]
fn child_frame(
    comments: &mut Vec<Comment>,
    reply_count: &mut usize,
    parent: &mut Frame,
    element: &BytesStart<'_>,
    decoder: Decoder,
    namespace: &str,
    local: &str,
    start: usize,
) -> Result<FrameKind> {
    match &mut parent.kind {
        FrameKind::Root => {
            if namespace != P188 || local != "cm" {
                return Err(invalid("cmLst permits only p188:cm children"));
            }
            if comments.len() >= MAX_COMMENTS {
                return Err(limit("modern comments"));
            }
            let declarations = namespace_declarations_from(element, decoder, None)?;
            let attributes = known_attributes(
                element,
                decoder,
                &[
                    "id",
                    "authorId",
                    "status",
                    "created",
                    "startDate",
                    "dueDate",
                    "assignedTo",
                    "complete",
                    "title",
                ],
            )?;
            comments.push(parse_comment_attributes(attributes, declarations)?);
            Ok(FrameKind::Comment {
                index: comments.len() - 1,
                stage: 0,
                anchor_family: None,
            })
        },
        FrameKind::Comment {
            index,
            stage,
            anchor_family,
        } => {
            let owner = RawOwner::Comment(*index);
            if let Some((family, kind)) = anchor_kind(namespace, local) {
                if *stage != 0 {
                    return Err(invalid("modern comment anchor is out of order"));
                }
                match (*anchor_family, family) {
                    (None, _) => *anchor_family = Some(family),
                    (Some(AnchorFamily::Drawing), AnchorFamily::Drawing)
                    | (Some(AnchorFamily::Text), AnchorFamily::Text) => {},
                    _ => return Err(invalid("modern comment mixes anchor choice branches")),
                }
                if matches!(family, AnchorFamily::Slide | AnchorFamily::Unknown)
                    && comments[*index]
                        .anchors
                        .iter()
                        .any(|anchor| anchor.kind == kind)
                {
                    return Err(invalid("singleton modern comment anchor is duplicated"));
                }
                if family == AnchorFamily::Unknown {
                    no_non_namespace_attributes(element)?;
                } else {
                    validate_any_attributes(element, decoder)?;
                }
                Ok(FrameKind::Raw {
                    start,
                    kind: RawKind::Anchor(*index, kind),
                })
            } else if namespace == P188 && local == "pos" {
                if *stage > 1 {
                    return Err(invalid("modern comment pos is duplicated or out of order"));
                }
                *stage = 2;
                let attributes = known_attributes(element, decoder, &["x", "y"])?;
                let x = required(&attributes, "x")?
                    .parse::<i64>()
                    .map_err(|_err| invalid("invalid modern comment x coordinate"))?;
                let y = required(&attributes, "y")?
                    .parse::<i64>()
                    .map_err(|_err| invalid("invalid modern comment y coordinate"))?;
                comments[*index].position = Some(Position { x, y });
                Ok(FrameKind::Position)
            } else if namespace == P188 && local == "replyLst" {
                if *stage > 2 || comments[*index].reply_list_present {
                    return Err(invalid(
                        "modern comment replyLst is duplicated or out of order",
                    ));
                }
                *stage = 3;
                no_non_namespace_attributes(element)?;
                comments[*index].reply_list_namespace_declarations =
                    namespace_declarations_from(element, decoder, None)?;
                comments[*index].reply_list_present = true;
                Ok(FrameKind::ReplyList { comment: *index })
            } else if namespace == P188 && local == "txBody" {
                if *stage > 3 || comments[*index].text_body_xml.is_some() {
                    return Err(invalid(
                        "modern comment txBody is duplicated or out of order",
                    ));
                }
                *stage = 4;
                validate_any_attributes(element, decoder)?;
                Ok(FrameKind::Raw {
                    start,
                    kind: RawKind::TextBody(owner),
                })
            } else if namespace == P188 && local == "extLst" {
                if *stage > 4 || comments[*index].extension_xml.is_some() {
                    return Err(invalid(
                        "modern comment extLst is duplicated or out of order",
                    ));
                }
                *stage = 5;
                validate_any_attributes(element, decoder)?;
                Ok(FrameKind::Raw {
                    start,
                    kind: RawKind::Extension(owner),
                })
            } else {
                Err(invalid("unexpected modern comment child"))
            }
        },
        FrameKind::ReplyList { comment } => {
            if namespace != P188 || local != "reply" {
                return Err(invalid("replyLst permits only p188:reply children"));
            }
            *reply_count = reply_count
                .checked_add(1)
                .ok_or_else(|| limit("modern comment replies"))?;
            if *reply_count > MAX_REPLIES {
                return Err(limit("modern comment replies"));
            }
            let declarations = namespace_declarations_from(element, decoder, None)?;
            let attributes =
                known_attributes(element, decoder, &["id", "authorId", "status", "created"])?;
            comments[*comment]
                .replies
                .push(parse_reply_attributes(attributes, declarations)?);
            Ok(FrameKind::Reply {
                comment: *comment,
                reply: comments[*comment].replies.len() - 1,
                stage: 0,
            })
        },
        FrameKind::Reply {
            comment,
            reply,
            stage,
        } => {
            let owner = RawOwner::Reply(*comment, *reply);
            if namespace == P188 && local == "txBody" {
                if *stage > 0 || comments[*comment].replies[*reply].text_body_xml.is_some() {
                    return Err(invalid(
                        "modern comment reply txBody is duplicated or out of order",
                    ));
                }
                *stage = 1;
                validate_any_attributes(element, decoder)?;
                Ok(FrameKind::Raw {
                    start,
                    kind: RawKind::TextBody(owner),
                })
            } else if namespace == P188 && local == "extLst" {
                if *stage > 1 || comments[*comment].replies[*reply].extension_xml.is_some() {
                    return Err(invalid(
                        "modern comment reply extLst is duplicated or out of order",
                    ));
                }
                *stage = 2;
                validate_any_attributes(element, decoder)?;
                Ok(FrameKind::Raw {
                    start,
                    kind: RawKind::Extension(owner),
                })
            } else {
                Err(invalid("unexpected modern comment reply child"))
            }
        },
        FrameKind::Raw { .. } | FrameKind::Opaque => {
            validate_any_attributes(element, decoder)?;
            Ok(FrameKind::Opaque)
        },
        FrameKind::Position => Err(invalid("modern comment pos must be empty")),
    }
}

fn parse_comment_attributes(
    attributes: HashMap<String, String>,
    namespace_declarations: Vec<NamespaceDeclaration>,
) -> Result<Comment> {
    let id = required(&attributes, "id")?.to_owned();
    let author_id = required(&attributes, "authorId")?.to_owned();
    let created = required(&attributes, "created")?.to_owned();
    validate_guid(&id)?;
    validate_guid(&author_id)?;
    validate_date_time(&created)?;
    let status = attributes
        .get("status")
        .map(|value| Status::parse(value))
        .transpose()?;
    let start_date = attributes.get("startDate").cloned();
    let due_date = attributes.get("dueDate").cloned();
    if let Some(value) = &start_date {
        validate_date_time(value)?;
    }
    if let Some(value) = &due_date {
        validate_date_time(value)?;
    }
    let assigned_to = attributes
        .get("assignedTo")
        .map(|value| {
            value
                .split_whitespace()
                .map(|id| {
                    validate_guid(id)?;
                    Ok(id.to_owned())
                })
                .collect::<Result<Vec<_>>>()
        })
        .transpose()?;
    let complete = attributes
        .get("complete")
        .map(|value| value.parse())
        .transpose()?;
    Ok(Comment {
        id,
        author_id,
        status,
        created,
        start_date,
        due_date,
        assigned_to,
        complete,
        title: attributes.get("title").cloned(),
        namespace_declarations,
        anchors: Vec::new(),
        position: None,
        reply_list_namespace_declarations: Vec::new(),
        replies: Vec::new(),
        reply_list_present: false,
        text_body_xml: None,
        extension_xml: None,
    })
}

fn parse_reply_attributes(
    attributes: HashMap<String, String>,
    namespace_declarations: Vec<NamespaceDeclaration>,
) -> Result<Reply> {
    let id = required(&attributes, "id")?.to_owned();
    let author_id = required(&attributes, "authorId")?.to_owned();
    let created = required(&attributes, "created")?.to_owned();
    validate_guid(&id)?;
    validate_guid(&author_id)?;
    validate_date_time(&created)?;
    Ok(Reply {
        id,
        author_id,
        status: attributes
            .get("status")
            .map(|value| Status::parse(value))
            .transpose()?,
        created,
        namespace_declarations,
        text_body_xml: None,
        extension_xml: None,
    })
}

fn anchor_kind(namespace: &str, local: &str) -> Option<(AnchorFamily, AnchorKind)> {
    match (namespace, local) {
        (PC, "sldMkLst") => Some((AnchorFamily::Slide, AnchorKind::SlideMoniker)),
        (AC, "deMkLst") => Some((AnchorFamily::Drawing, AnchorKind::DrawingElementMoniker)),
        (AC, "txMkLst") => Some((AnchorFamily::Text, AnchorKind::TextRangeMoniker)),
        (P188, "unknownAnchor") => Some((AnchorFamily::Unknown, AnchorKind::Unknown)),
        _ => None,
    }
}

fn close_empty(kind: &FrameKind) -> Result<()> {
    match kind {
        FrameKind::Position
        | FrameKind::Root
        | FrameKind::Comment { .. }
        | FrameKind::ReplyList { .. }
        | FrameKind::Reply { .. }
        | FrameKind::Raw { .. }
        | FrameKind::Opaque => Ok(()),
    }
}

fn attach_raw(
    kind: &FrameKind,
    bytes: &[u8],
    empty_start: usize,
    end: usize,
    comments: &mut [Comment],
) -> Result<()> {
    let FrameKind::Raw { start, kind } = kind else {
        return Ok(());
    };
    let start = if empty_start == 0 {
        *start
    } else {
        empty_start
    };
    if start > end || end > bytes.len() {
        return Err(invalid("invalid modern Comment XML fragment bounds"));
    }
    let xml = bytes[start..end].to_vec();
    match *kind {
        RawKind::Anchor(comment, anchor_kind) => comments[comment].anchors.push(Anchor {
            kind: anchor_kind,
            xml,
        }),
        RawKind::TextBody(RawOwner::Comment(comment)) => {
            comments[comment].text_body_xml = Some(xml);
        },
        RawKind::Extension(RawOwner::Comment(comment)) => {
            comments[comment].extension_xml = Some(xml);
        },
        RawKind::TextBody(RawOwner::Reply(comment, reply)) => {
            comments[comment].replies[reply].text_body_xml = Some(xml);
        },
        RawKind::Extension(RawOwner::Reply(comment, reply)) => {
            comments[comment].replies[reply].extension_xml = Some(xml);
        },
    }
    Ok(())
}

fn inside_raw(stack: &[Frame]) -> bool {
    stack
        .iter()
        .any(|frame| matches!(frame.kind, FrameKind::Raw { .. }))
}

fn required<'a>(attributes: &'a HashMap<String, String>, name: &str) -> Result<&'a str> {
    attributes
        .get(name)
        .map(String::as_str)
        .ok_or_else(|| invalid(format!("modern comment is missing required '{name}'")))
}

fn known_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&str],
) -> Result<HashMap<String, String>> {
    let mut values = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = decode_name(attribute.key.as_ref())?;
        if is_namespace_attribute(&key) {
            continue;
        }
        if key.contains(':') || !allowed.contains(&key.as_str()) {
            return Err(invalid(format!("unexpected attribute '{key}'")));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded(&value)?;
        if values.insert(key.clone(), value).is_some() {
            return Err(invalid(format!("duplicate attribute '{key}'")));
        }
    }
    Ok(values)
}

fn validate_any_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        bounded(&value)?;
    }
    Ok(())
}

fn no_non_namespace_attributes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if !is_namespace_attribute(&decode_name(attribute.key.as_ref())?) {
            return Err(invalid("unexpected attribute on modern Comment container"));
        }
    }
    Ok(())
}

fn namespace_declarations_from(
    element: &BytesStart<'_>,
    decoder: Decoder,
    exclude_prefix: Option<&str>,
) -> Result<Vec<NamespaceDeclaration>> {
    let mut result = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = decode_name(attribute.key.as_ref())?;
        let prefix = if key == "xmlns" {
            Some(String::new())
        } else {
            key.strip_prefix("xmlns:").map(str::to_owned)
        };
        let Some(prefix) = prefix else {
            continue;
        };
        if exclude_prefix == Some(prefix.as_str()) {
            continue;
        }
        let uri = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        result.push(NamespaceDeclaration { prefix, uri });
    }
    validate_namespaces(&result, None)?;
    Ok(result)
}

fn resolve_namespace(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(value) => Ok(std::str::from_utf8(value.as_ref())
            .map_err(xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn element_prefix(element: &BytesStart<'_>) -> Result<String> {
    let name = decode_name(element.name().as_ref())?;
    Ok(name
        .rsplit_once(':')
        .map_or(String::new(), |(prefix, _)| prefix.to_owned()))
}

fn decode_name(value: &[u8]) -> Result<String> {
    Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
}

fn is_namespace_attribute(value: &str) -> bool {
    value == "xmlns" || value.starts_with("xmlns:")
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
