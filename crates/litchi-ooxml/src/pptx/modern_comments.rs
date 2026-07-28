//! PowerPoint 2018 modern Comment parts.
//!
//! Imported moniker, DrawingML text-body, and extension payloads are retained
//! as XML and are never interpreted or used to resolve relationships.

use crate::common::mce::process_ooxml;
use crate::error::{OoxmlError, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

pub const MODERN_COMMENT_CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.comments+xml";
pub const MODERN_COMMENT_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/comments";

const P188: &str = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
const PC: &str = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
const AC: &str = "http://schemas.microsoft.com/office/drawing/2013/main/command";
const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const MAX_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 192;
const MAX_NODES: usize = 250_000;
const MAX_COMMENTS: usize = 100_000;
const MAX_REPLIES: usize = 100_000;
const MAX_STRING_BYTES: usize = 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModernCommentNamespaceDeclaration {
    /// Empty means the default namespace.
    pub prefix: String,
    pub uri: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ModernCommentStatus {
    Active,
    Resolved,
    Closed,
}

impl ModernCommentStatus {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "active" => Ok(Self::Active),
            "resolved" => Ok(Self::Resolved),
            "closed" => Ok(Self::Closed),
            _ => Err(invalid(format!("invalid modern comment status '{value}'"))),
        }
    }

    fn token(self) -> &'static str {
        match self {
            Self::Active => "active",
            Self::Resolved => "resolved",
            Self::Closed => "closed",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ModernCommentAnchorKind {
    SlideMoniker,
    DrawingElementMoniker,
    TextRangeMoniker,
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModernCommentAnchor {
    pub kind: ModernCommentAnchorKind,
    /// Complete imported moniker element retained inertly.
    pub xml: Vec<u8>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ModernCommentPosition {
    pub x: i64,
    pub y: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModernCommentReply {
    pub id: String,
    pub author_id: String,
    /// `None` retains omission and the schema default `active`.
    pub status: Option<ModernCommentStatus>,
    pub created: String,
    pub namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
    /// Optional complete `p188:txBody` fragment retained inertly.
    pub text_body_xml: Option<Vec<u8>>,
    /// Optional complete `p188:extLst` fragment retained inertly.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModernComment {
    pub id: String,
    pub author_id: String,
    /// `None` retains omission and the schema default `active`.
    pub status: Option<ModernCommentStatus>,
    pub created: String,
    pub start_date: Option<String>,
    pub due_date: Option<String>,
    /// `None` distinguishes omission from a present empty list.
    pub assigned_to: Option<Vec<String>>,
    /// Original valid `ST_PositiveFixedPercentage` lexical value.
    pub complete: Option<String>,
    pub title: Option<String>,
    pub namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
    pub anchors: Vec<ModernCommentAnchor>,
    pub position: Option<ModernCommentPosition>,
    pub reply_list_namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
    pub replies: Vec<ModernCommentReply>,
    /// Whether the optional `replyLst` wrapper was present, including when empty.
    pub reply_list_present: bool,
    /// Optional complete `p188:txBody` fragment retained inertly.
    pub text_body_xml: Option<Vec<u8>>,
    /// Optional complete `p188:extLst` fragment retained inertly.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ModernCommentList {
    /// Prefix used for the 2018 PowerPoint namespace. Empty means default.
    pub root_prefix: String,
    pub namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
    pub comments: Vec<ModernComment>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModernCommentPart {
    pub slide_part_name: String,
    pub relationship_id: String,
    pub part_name: String,
    pub comments: ModernCommentList,
}

impl ModernCommentList {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        parse_comment_list(xml)
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        validate_model(self)?;
        let mut out = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
        open_tag(&mut out, &self.root_prefix, "cmLst");
        write_namespace_binding(&mut out, &self.root_prefix, P188);
        write_namespaces(&mut out, &self.namespace_declarations);
        if self.comments.is_empty() {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            for comment in &self.comments {
                write_comment(&mut out, &self.root_prefix, comment);
            }
            close_tag(&mut out, &self.root_prefix, "cmLst");
        }
        if out.len() > MAX_BYTES {
            return Err(limit("serialized modern Comment bytes"));
        }
        parse_comment_list(&out)?;
        Ok(out)
    }
}

pub fn load_modern_comments(package: &OpcPackage) -> Result<Vec<ModernCommentPart>> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == MODERN_COMMENT_RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "modern Comment relationship cannot originate at the package root",
        ));
    }

    let mut relationships = Vec::new();
    let mut targets = HashSet::new();
    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == MODERN_COMMENT_RELATIONSHIP_TYPE)
        {
            if source.content_type() != SLIDE_CONTENT_TYPE {
                return Err(invalid(format!(
                    "modern Comment relationship has non-Slide source '{}'",
                    source.partname()
                )));
            }
            if relationship.is_external() {
                return Err(invalid("modern Comment relationship cannot be external"));
            }
            let target = relationship.target_partname()?;
            let part = package.get_part(&target)?;
            require_content_type(part, MODERN_COMMENT_CONTENT_TYPE)?;
            targets.insert(target.to_string());
            relationships.push((
                source.partname().to_string(),
                relationship.r_id().to_string(),
                target.to_string(),
            ));
        }
    }

    for part in package.iter_parts() {
        if part.content_type() == MODERN_COMMENT_CONTENT_TYPE
            && !targets.contains(part.partname().as_str())
        {
            return Err(invalid(format!(
                "package contains orphan modern Comment part '{}'",
                part.partname()
            )));
        }
    }

    relationships
        .into_iter()
        .map(|(slide_part_name, relationship_id, part_name)| {
            let uri = PackURI::new(&part_name).map_err(OoxmlError::InvalidUri)?;
            let comments = ModernCommentList::parse(package.get_part(&uri)?.blob())?;
            Ok(ModernCommentPart {
                slide_part_name,
                relationship_id,
                part_name,
                comments,
            })
        })
        .collect()
}

/// Add a new modern Comment part after validating the complete existing graph.
/// Existing parts are deliberately not overwritten.
pub fn store_modern_comment(package: &mut OpcPackage, value: &ModernCommentPart) -> Result<()> {
    load_modern_comments(package)?;
    validate_relationship_id(&value.relationship_id)?;
    let slide_name = PackURI::new(&value.slide_part_name).map_err(OoxmlError::InvalidUri)?;
    let slide = package.get_part(&slide_name)?;
    if slide.content_type() != SLIDE_CONTENT_TYPE {
        return Err(invalid(format!(
            "'{}' is not a Slide part",
            value.slide_part_name
        )));
    }
    if slide.rels().get(&value.relationship_id).is_some() {
        return Err(invalid("modern Comment relationship ID already exists"));
    }
    let part_name = PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    if package
        .iter_parts()
        .any(|part| part.partname() == &part_name)
    {
        return Err(invalid(format!("part '{part_name}' already exists")));
    }
    let xml = value.comments.to_xml()?;
    let target = part_name.relative_ref(slide_name.base_uri());
    package.try_add_part(Box::new(BlobPart::new(
        part_name,
        MODERN_COMMENT_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(&slide_name)?
        .rels_mut()
        .add_relationship(
            MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
            target,
            value.relationship_id.clone(),
            false,
        );
    Ok(())
}

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
    Anchor(usize, ModernCommentAnchorKind),
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

fn parse_comment_list(xml: &[u8]) -> Result<ModernCommentList> {
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
                    child_frame(
                        &mut comments,
                        &mut reply_count,
                        stack.last_mut().expect("nonempty stack"),
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
    let value = ModernCommentList {
        root_prefix,
        namespace_declarations,
        comments,
    };
    validate_model(&value)?;
    Ok(value)
}

#[allow(clippy::too_many_arguments)]
fn child_frame(
    comments: &mut Vec<ModernComment>,
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
                    .map_err(|_| invalid("invalid modern comment x coordinate"))?;
                let y = required(&attributes, "y")?
                    .parse::<i64>()
                    .map_err(|_| invalid("invalid modern comment y coordinate"))?;
                comments[*index].position = Some(ModernCommentPosition { x, y });
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
    namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
) -> Result<ModernComment> {
    let id = required(&attributes, "id")?.to_owned();
    let author_id = required(&attributes, "authorId")?.to_owned();
    let created = required(&attributes, "created")?.to_owned();
    validate_guid(&id)?;
    validate_guid(&author_id)?;
    validate_date_time(&created)?;
    let status = attributes
        .get("status")
        .map(|value| ModernCommentStatus::parse(value))
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
    let complete = attributes.get("complete").cloned();
    if let Some(value) = &complete {
        validate_percentage(value)?;
    }
    Ok(ModernComment {
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
    namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
) -> Result<ModernCommentReply> {
    let id = required(&attributes, "id")?.to_owned();
    let author_id = required(&attributes, "authorId")?.to_owned();
    let created = required(&attributes, "created")?.to_owned();
    validate_guid(&id)?;
    validate_guid(&author_id)?;
    validate_date_time(&created)?;
    Ok(ModernCommentReply {
        id,
        author_id,
        status: attributes
            .get("status")
            .map(|value| ModernCommentStatus::parse(value))
            .transpose()?,
        created,
        namespace_declarations,
        text_body_xml: None,
        extension_xml: None,
    })
}

fn anchor_kind(namespace: &str, local: &str) -> Option<(AnchorFamily, ModernCommentAnchorKind)> {
    match (namespace, local) {
        (PC, "sldMkLst") => Some((AnchorFamily::Slide, ModernCommentAnchorKind::SlideMoniker)),
        (AC, "deMkLst") => Some((
            AnchorFamily::Drawing,
            ModernCommentAnchorKind::DrawingElementMoniker,
        )),
        (AC, "txMkLst") => Some((
            AnchorFamily::Text,
            ModernCommentAnchorKind::TextRangeMoniker,
        )),
        (P188, "unknownAnchor") => Some((AnchorFamily::Unknown, ModernCommentAnchorKind::Unknown)),
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
    comments: &mut [ModernComment],
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
        RawKind::Anchor(comment, anchor_kind) => {
            comments[comment].anchors.push(ModernCommentAnchor {
                kind: anchor_kind,
                xml,
            })
        },
        RawKind::TextBody(RawOwner::Comment(comment)) => {
            comments[comment].text_body_xml = Some(xml)
        },
        RawKind::Extension(RawOwner::Comment(comment)) => {
            comments[comment].extension_xml = Some(xml)
        },
        RawKind::TextBody(RawOwner::Reply(comment, reply)) => {
            comments[comment].replies[reply].text_body_xml = Some(xml)
        },
        RawKind::Extension(RawOwner::Reply(comment, reply)) => {
            comments[comment].replies[reply].extension_xml = Some(xml)
        },
    }
    Ok(())
}

fn inside_raw(stack: &[Frame]) -> bool {
    stack
        .iter()
        .any(|frame| matches!(frame.kind, FrameKind::Raw { .. }))
}

fn validate_model(value: &ModernCommentList) -> Result<()> {
    validate_prefix(&value.root_prefix)?;
    validate_namespaces(&value.namespace_declarations, Some(&value.root_prefix))?;
    if value.comments.len() > MAX_COMMENTS {
        return Err(limit("modern comments"));
    }
    let mut replies = 0usize;
    for comment in &value.comments {
        validate_guid(&comment.id)?;
        validate_guid(&comment.author_id)?;
        validate_date_time(&comment.created)?;
        if let Some(value) = &comment.start_date {
            validate_date_time(value)?;
        }
        if let Some(value) = &comment.due_date {
            validate_date_time(value)?;
        }
        if let Some(ids) = &comment.assigned_to {
            for id in ids {
                validate_guid(id)?;
            }
        }
        if let Some(value) = &comment.complete {
            validate_percentage(value)?;
        }
        validate_optional_string(comment.title.as_deref())?;
        validate_namespaces(&comment.namespace_declarations, None)?;
        validate_namespaces(&comment.reply_list_namespace_declarations, None)?;
        if !comment.reply_list_present && !comment.replies.is_empty() {
            return Err(invalid("modern comment replies require replyLst presence"));
        }
        replies = replies
            .checked_add(comment.replies.len())
            .ok_or_else(|| limit("modern comment replies"))?;
        if replies > MAX_REPLIES {
            return Err(limit("modern comment replies"));
        }
        for anchor in &comment.anchors {
            if anchor.xml.len() > MAX_BYTES {
                return Err(limit("modern comment anchor bytes"));
            }
        }
        validate_fragment(comment.text_body_xml.as_deref())?;
        validate_fragment(comment.extension_xml.as_deref())?;
        for reply in &comment.replies {
            validate_guid(&reply.id)?;
            validate_guid(&reply.author_id)?;
            validate_date_time(&reply.created)?;
            validate_namespaces(&reply.namespace_declarations, None)?;
            validate_fragment(reply.text_body_xml.as_deref())?;
            validate_fragment(reply.extension_xml.as_deref())?;
        }
    }
    Ok(())
}

fn write_comment(out: &mut Vec<u8>, prefix: &str, comment: &ModernComment) {
    open_tag(out, prefix, "cm");
    write_attr(out, "id", &comment.id);
    write_attr(out, "authorId", &comment.author_id);
    if let Some(status) = comment.status {
        write_attr(out, "status", status.token());
    }
    write_attr(out, "created", &comment.created);
    if let Some(value) = &comment.start_date {
        write_attr(out, "startDate", value);
    }
    if let Some(value) = &comment.due_date {
        write_attr(out, "dueDate", value);
    }
    if let Some(values) = &comment.assigned_to {
        write_attr(out, "assignedTo", &values.join(" "));
    }
    if let Some(value) = &comment.complete {
        write_attr(out, "complete", value);
    }
    if let Some(value) = &comment.title {
        write_attr(out, "title", value);
    }
    write_namespaces(out, &comment.namespace_declarations);
    let has_children = !comment.anchors.is_empty()
        || comment.position.is_some()
        || comment.reply_list_present
        || comment.text_body_xml.is_some()
        || comment.extension_xml.is_some();
    if !has_children {
        out.extend_from_slice(b"/>");
        return;
    }
    out.push(b'>');
    for anchor in &comment.anchors {
        out.extend_from_slice(&anchor.xml);
    }
    if let Some(position) = comment.position {
        open_tag(out, prefix, "pos");
        write_attr(out, "x", &position.x.to_string());
        write_attr(out, "y", &position.y.to_string());
        out.extend_from_slice(b"/>");
    }
    if comment.reply_list_present {
        open_tag(out, prefix, "replyLst");
        write_namespaces(out, &comment.reply_list_namespace_declarations);
        if comment.replies.is_empty() {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            for reply in &comment.replies {
                write_reply(out, prefix, reply);
            }
            close_tag(out, prefix, "replyLst");
        }
    }
    if let Some(xml) = &comment.text_body_xml {
        out.extend_from_slice(xml);
    }
    if let Some(xml) = &comment.extension_xml {
        out.extend_from_slice(xml);
    }
    close_tag(out, prefix, "cm");
}

fn write_reply(out: &mut Vec<u8>, prefix: &str, reply: &ModernCommentReply) {
    open_tag(out, prefix, "reply");
    write_attr(out, "id", &reply.id);
    write_attr(out, "authorId", &reply.author_id);
    if let Some(status) = reply.status {
        write_attr(out, "status", status.token());
    }
    write_attr(out, "created", &reply.created);
    write_namespaces(out, &reply.namespace_declarations);
    if reply.text_body_xml.is_none() && reply.extension_xml.is_none() {
        out.extend_from_slice(b"/>");
    } else {
        out.push(b'>');
        if let Some(xml) = &reply.text_body_xml {
            out.extend_from_slice(xml);
        }
        if let Some(xml) = &reply.extension_xml {
            out.extend_from_slice(xml);
        }
        close_tag(out, prefix, "reply");
    }
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
) -> Result<Vec<ModernCommentNamespaceDeclaration>> {
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
        result.push(ModernCommentNamespaceDeclaration { prefix, uri });
    }
    validate_namespaces(&result, None)?;
    Ok(result)
}

fn validate_namespaces(
    value: &[ModernCommentNamespaceDeclaration],
    excluded: Option<&str>,
) -> Result<()> {
    let mut seen = HashSet::new();
    for declaration in value {
        validate_prefix(&declaration.prefix)?;
        bounded(&declaration.uri)?;
        if declaration.prefix == "xml" || declaration.prefix == "xmlns" {
            return Err(invalid(
                "reserved XML namespace prefix cannot be redeclared",
            ));
        }
        if excluded == Some(declaration.prefix.as_str()) {
            return Err(invalid("modern Comment namespace prefix is declared twice"));
        }
        if !seen.insert(&declaration.prefix) {
            return Err(invalid("duplicate namespace declaration"));
        }
    }
    Ok(())
}

fn validate_prefix(value: &str) -> Result<()> {
    if value.is_empty()
        || (value.bytes().enumerate().all(|(index, byte)| {
            byte.is_ascii_alphanumeric()
                || byte == b'_'
                || byte == b'-'
                || (byte == b'.' && index != 0)
        }) && !value.as_bytes()[0].is_ascii_digit()
            && !value.starts_with('-'))
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid XML namespace prefix '{value}'")))
    }
}

fn validate_guid(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    if bytes.len() == 38
        && bytes[0] == b'{'
        && bytes[37] == b'}'
        && [9, 14, 19, 24].iter().all(|index| bytes[*index] == b'-')
        && bytes[1..37]
            .iter()
            .enumerate()
            .all(|(index, byte)| [8, 13, 18, 23].contains(&index) || byte.is_ascii_hexdigit())
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid modern Comment GUID '{value}'")))
    }
}

fn validate_date_time(value: &str) -> Result<()> {
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid XML dateTime '{value}'")))
    }
}

fn validate_percentage(value: &str) -> Result<()> {
    let valid = if let Some(number) = value.strip_suffix('%') {
        !number.is_empty()
            && !number.contains(['e', 'E'])
            && number
                .parse::<f64>()
                .is_ok_and(|number| number.is_finite() && (0.0..=100.0).contains(&number))
    } else {
        value.parse::<u32>().is_ok_and(|number| number <= 100_000)
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(format!(
            "invalid positive fixed percentage '{value}'"
        )))
    }
}

fn validate_optional_string(value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        bounded(value)?;
    }
    Ok(())
}

fn validate_fragment(value: Option<&[u8]>) -> Result<()> {
    if value.is_some_and(|value| value.len() > MAX_BYTES) {
        Err(limit("modern Comment opaque fragment bytes"))
    } else {
        Ok(())
    }
}

fn validate_relationship_id(value: &str) -> Result<()> {
    bounded(value)?;
    if value.is_empty() || value.chars().any(char::is_whitespace) {
        Err(invalid(
            "modern Comment relationship ID must be nonempty without whitespace",
        ))
    } else {
        Ok(())
    }
}

fn required<'a>(attributes: &'a HashMap<String, String>, name: &str) -> Result<&'a str> {
    attributes
        .get(name)
        .map(String::as_str)
        .ok_or_else(|| invalid(format!("modern comment is missing required '{name}'")))
}

fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("modern Comment string bytes"))
    }
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

fn open_tag(out: &mut Vec<u8>, prefix: &str, local: &str) {
    out.push(b'<');
    qname(out, prefix, local);
}

fn close_tag(out: &mut Vec<u8>, prefix: &str, local: &str) {
    out.extend_from_slice(b"</");
    qname(out, prefix, local);
    out.push(b'>');
}

fn qname(out: &mut Vec<u8>, prefix: &str, local: &str) {
    if !prefix.is_empty() {
        out.extend_from_slice(prefix.as_bytes());
        out.push(b':');
    }
    out.extend_from_slice(local.as_bytes());
}

fn write_namespace_binding(out: &mut Vec<u8>, prefix: &str, uri: &str) {
    out.extend_from_slice(b" xmlns");
    if !prefix.is_empty() {
        out.push(b':');
        out.extend_from_slice(prefix.as_bytes());
    }
    out.extend_from_slice(b"=\"");
    escape(out, uri);
    out.push(b'"');
}

fn write_namespaces(out: &mut Vec<u8>, declarations: &[ModernCommentNamespaceDeclaration]) {
    for declaration in declarations {
        write_namespace_binding(out, &declaration.prefix, &declaration.uri);
    }
}

fn write_attr(out: &mut Vec<u8>, name: &str, value: &str) {
    out.push(b' ');
    out.extend_from_slice(name.as_bytes());
    out.extend_from_slice(b"=\"");
    escape(out, value);
    out.push(b'"');
}

fn escape(out: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                out.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(OoxmlError::InvalidContentType {
            expected: expected.into(),
            got: part.content_type().into(),
        })
    }
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(label: &str) -> OoxmlError {
    invalid(format!("{label} exceeds implementation limit"))
}

/// Find a modern comment by slide part and stable comment GUID.
pub fn find_modern_comment(
    package: &OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
) -> Result<Option<ModernComment>> {
    Ok(load_modern_comments(package)?
        .into_iter()
        .find(|part| part.slide_part_name == slide_part_name.to_string())
        .and_then(|part| part.comments.comments.into_iter().find(|comment| comment.id == comment_id)))
}

/// Add a modern comment to a slide, creating a collision-safe part when necessary.
pub fn add_modern_comment(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    comment: ModernComment,
) -> Result<ModernCommentPart> {
    let mut parts = load_modern_comments(package)?;
    ensure_modern_comment_id_is_free(&parts, &comment.id)?;
    ensure_modern_reply_ids_are_free(&parts, &comment)?;
    if let Some(index) = parts
        .iter()
        .position(|part| part.slide_part_name == slide_part_name.to_string())
    {
        let mut staged = parts[index].clone();
        staged.comments.comments.push(comment);
        parts[index] = staged.clone();
        validate_and_commit_modern_comment_part(package, &staged, &parts)?;
        return Ok(staged);
    }

    package.get_part(slide_part_name)?;
    let part_name = next_modern_comment_part_name(package)?;
    let relationship_id = next_modern_comment_relationship_id(package, slide_part_name)?;
    let staged = ModernCommentPart {
        slide_part_name: slide_part_name.to_string(),
        relationship_id: relationship_id.clone(),
        part_name: part_name.to_string(),
        comments: ModernCommentList {
            root_prefix: "p188".into(),
            namespace_declarations: Vec::new(),
            comments: vec![comment],
        },
    };
    parts.push(staged.clone());
    validate_modern_comment_graph_for_mutation(package, &parts)?;
    let xml = staged.comments.to_xml()?;
    ModernCommentList::parse(&xml)?;
    package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
        part_name.clone(),
        MODERN_COMMENT_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(slide_part_name)?
        .rels_mut()
        .add_relationship(
            MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
            part_name.relative_ref(slide_part_name.base_uri()),
            relationship_id,
            false,
        );
    let _ = package.clear_digital_signatures();
    Ok(staged)
}

/// Update a modern comment without permitting its stable GUID to change.
pub fn update_modern_comment<F>(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
    update: F,
) -> Result<bool>
where
    F: FnOnce(&mut ModernComment),
{
    let mut parts = load_modern_comments(package)?;
    let Some(part_index) = parts
        .iter()
        .position(|part| part.slide_part_name == slide_part_name.to_string())
    else { return Ok(false); };
    let Some(comment_index) = parts[part_index]
        .comments
        .comments
        .iter()
        .position(|comment| comment.id == comment_id)
    else { return Ok(false); };
    let reply_ids_before: Vec<String> = parts[part_index].comments.comments[comment_index]
        .replies
        .iter()
        .map(|reply| reply.id.clone())
        .collect();
    update(&mut parts[part_index].comments.comments[comment_index]);
    if parts[part_index].comments.comments[comment_index].id != comment_id {
        return Err(invalid("modern comment update cannot change its ID"));
    }
    let reply_ids_after: Vec<&str> = parts[part_index].comments.comments[comment_index]
        .replies
        .iter()
        .map(|reply| reply.id.as_str())
        .collect();
    if reply_ids_after.len() == reply_ids_before.len()
        && reply_ids_after
            .iter()
            .zip(&reply_ids_before)
            .any(|(after, before)| *after != before)
    {
        return Err(invalid("modern comment update cannot change a reply ID"));
    }
    ensure_all_modern_ids_are_unique(&parts)?;
    let staged = parts[part_index].clone();
    validate_and_commit_modern_comment_part(package, &staged, &parts)?;
    Ok(true)
}

/// Replace a modern comment without permitting its stable GUID to change.
pub fn replace_modern_comment(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
    replacement: ModernComment,
) -> Result<bool> {
    if replacement.id != comment_id {
        return Err(invalid("replacement modern comment ID must match"));
    }
    update_modern_comment(package, slide_part_name, comment_id, move |comment| {
        *comment = replacement;
    })
}

/// Remove a modern comment and remove an empty per-slide part unless another owner shares it.
pub fn remove_modern_comment(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
) -> Result<bool> {
    let mut parts = load_modern_comments(package)?;
    let Some(part_index) = parts
        .iter()
        .position(|part| part.slide_part_name == slide_part_name.to_string())
    else { return Ok(false); };
    let Some(comment_index) = parts[part_index]
        .comments
        .comments
        .iter()
        .position(|comment| comment.id == comment_id)
    else { return Ok(false); };
    parts[part_index].comments.comments.remove(comment_index);
    if !parts[part_index].comments.comments.is_empty() {
        let staged = parts[part_index].clone();
        validate_and_commit_modern_comment_part(package, &staged, &parts)?;
        return Ok(true);
    }

    let removed = parts.remove(part_index);
    validate_modern_comment_graph_for_mutation(package, &parts)?;
    let part_name = PackURI::new(&removed.part_name).map_err(invalid)?;
    package
        .get_part_mut(slide_part_name)?
        .rels_mut()
        .remove(&removed.relationship_id);
    if !modern_comment_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

/// Reorder every modern comment in one slide part by a complete GUID list.
pub fn reorder_modern_comments(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    ordered_comment_ids: &[String],
) -> Result<Vec<ModernComment>> {
    let mut parts = load_modern_comments(package)?;
    let Some(part_index) = parts
        .iter()
        .position(|part| part.slide_part_name == slide_part_name.to_string())
    else {
        if ordered_comment_ids.is_empty() { return Ok(Vec::new()); }
        return Err(invalid("modern comment part is missing for slide"));
    };
    if ordered_comment_ids.len() != parts[part_index].comments.comments.len() {
        return Err(invalid("modern comment reorder must contain every comment"));
    }
    let mut remaining = std::collections::HashMap::new();
    for comment in parts[part_index].comments.comments.drain(..) {
        if remaining.insert(comment.id.clone(), comment).is_some() {
            return Err(invalid("duplicate modern comment ID"));
        }
    }
    let mut ordered = Vec::with_capacity(ordered_comment_ids.len());
    for id in ordered_comment_ids {
        let comment = remaining
            .remove(id)
            .ok_or_else(|| invalid(format!("unknown or duplicate modern comment ID {id}")))?;
        ordered.push(comment);
    }
    if !remaining.is_empty() {
        return Err(invalid("modern comment reorder must contain every comment"));
    }
    parts[part_index].comments.comments = ordered.clone();
    let staged = parts[part_index].clone();
    validate_and_commit_modern_comment_part(package, &staged, &parts)?;
    Ok(ordered)
}

/// Find a reply by its stable GUID within a modern comment thread.
pub fn find_modern_comment_reply(
    package: &OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
    reply_id: &str,
) -> Result<Option<ModernCommentReply>> {
    Ok(find_modern_comment(package, slide_part_name, comment_id)?
        .and_then(|comment| comment.replies.into_iter().find(|reply| reply.id == reply_id)))
}

/// Add a reply to a modern comment thread.
pub fn add_modern_comment_reply(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
    reply: ModernCommentReply,
) -> Result<bool> {
    let reply_id = reply.id.clone();
    let parts = load_modern_comments(package)?;
    if modern_id_exists(&parts, &reply_id) {
        return Err(invalid(format!("duplicate modern comment or reply ID {reply_id}")));
    }
    update_modern_comment(package, slide_part_name, comment_id, move |comment| {
        comment.reply_list_present = true;
        comment.replies.push(reply);
    })
}

/// Update a reply without permitting its stable GUID to change.
pub fn update_modern_comment_reply<F>(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
    reply_id: &str,
    update: F,
) -> Result<bool>
where
    F: FnOnce(&mut ModernCommentReply),
{
    if find_modern_comment_reply(package, slide_part_name, comment_id, reply_id)?.is_none() {
        return Ok(false);
    }
    let mut update = Some(update);
    let mut found = false;
    let result = update_modern_comment(package, slide_part_name, comment_id, |comment| {
        if let Some(reply) = comment.replies.iter_mut().find(|reply| reply.id == reply_id) {
            if let Some(update) = update.take() { update(reply); }
            found = true;
        }
    })?;
    if !result || !found { return Ok(false); }
    Ok(true)
}

/// Replace a reply without permitting its stable GUID to change.
pub fn replace_modern_comment_reply(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
    reply_id: &str,
    replacement: ModernCommentReply,
) -> Result<bool> {
    if replacement.id != reply_id {
        return Err(invalid("replacement modern reply ID must match"));
    }
    update_modern_comment_reply(package, slide_part_name, comment_id, reply_id, move |reply| {
        *reply = replacement;
    })
}

/// Remove a reply from a modern comment thread.
pub fn remove_modern_comment_reply(
    package: &mut OpcPackage,
    slide_part_name: &PackURI,
    comment_id: &str,
    reply_id: &str,
) -> Result<bool> {
    if find_modern_comment_reply(package, slide_part_name, comment_id, reply_id)?.is_none() {
        return Ok(false);
    }
    let mut found = false;
    let result = update_modern_comment(package, slide_part_name, comment_id, |comment| {
        if let Some(index) = comment.replies.iter().position(|reply| reply.id == reply_id) {
            comment.replies.remove(index);
            found = true;
        }
    })?;
    Ok(result && found)
}

fn validate_and_commit_modern_comment_part(
    package: &mut OpcPackage,
    staged: &ModernCommentPart,
    all_parts: &[ModernCommentPart],
) -> Result<()> {
    validate_modern_comment_graph_for_mutation(package, all_parts)?;
    let xml = staged.comments.to_xml()?;
    ModernCommentList::parse(&xml)?;
    let part_name = PackURI::new(&staged.part_name).map_err(invalid)?;
    package.get_part_mut(&part_name)?.set_blob(xml);
    let _ = package.clear_digital_signatures();
    Ok(())
}

fn validate_modern_comment_graph_for_mutation(
    package: &OpcPackage,
    parts: &[ModernCommentPart],
) -> Result<()> {
    ensure_all_modern_ids_are_unique(parts)?;
    let authors = super::modern_comment_authors::load_modern_comment_authors(package)?;
    super::modern_comment_authors::validate_modern_comment_author_references(
        authors.as_ref(),
        parts,
    )
}

fn ensure_modern_comment_id_is_free(parts: &[ModernCommentPart], id: &str) -> Result<()> {
    if modern_id_exists(parts, id) {
        Err(invalid(format!("duplicate modern comment or reply ID {id}")))
    } else {
        Ok(())
    }
}

fn ensure_modern_reply_ids_are_free(parts: &[ModernCommentPart], comment: &ModernComment) -> Result<()> {
    let mut ids = std::collections::HashSet::new();
    for reply in &comment.replies {
        if reply.id == comment.id || !ids.insert(reply.id.clone()) || modern_id_exists(parts, &reply.id) {
            return Err(invalid(format!("duplicate modern comment or reply ID {}", reply.id)));
        }
    }
    Ok(())
}

fn ensure_all_modern_ids_are_unique(parts: &[ModernCommentPart]) -> Result<()> {
    let mut ids = std::collections::HashSet::new();
    for part in parts {
        for comment in &part.comments.comments {
            if !ids.insert(comment.id.clone()) {
                return Err(invalid(format!("duplicate modern comment or reply ID {}", comment.id)));
            }
            for reply in &comment.replies {
                if !ids.insert(reply.id.clone()) {
                    return Err(invalid(format!("duplicate modern comment or reply ID {}", reply.id)));
                }
            }
        }
    }
    Ok(())
}

fn modern_id_exists(parts: &[ModernCommentPart], id: &str) -> bool {
    parts.iter().any(|part| part.comments.comments.iter().any(|comment| {
        comment.id == id || comment.replies.iter().any(|reply| reply.id == id)
    }))
}

fn next_modern_comment_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 1..=65_537u32 {
        let candidate = PackURI::new(format!("/ppt/comments/modernComment{suffix}.xml"))
            .map_err(invalid)?;
        if package.get_part(&candidate).is_err() { return Ok(candidate); }
    }
    Err(invalid("no free modern comment part name"))
}

fn next_modern_comment_relationship_id(
    package: &OpcPackage,
    slide_part_name: &PackURI,
) -> Result<String> {
    let relationships = package.get_part(slide_part_name)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdModernComments{suffix}");
        if relationships.get(&candidate).is_none() { return Ok(candidate); }
    }
    Err(invalid("no free modern comment relationship ID"))
}

fn modern_comment_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| part.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship.target_partname().is_ok_and(|name| name == *target)
    })) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship.target_partname().is_ok_and(|name| name == *target)
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const AUTHOR: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
    const COMMENT: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";
    const REPLY: &str = "{E524A04C-CF22-45D7-A60D-09322EA5A80D}";

    fn sdk_xml() -> Vec<u8> {
        format!(r#"<p188:cmLst xmlns:p188="{P188}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Needs more cowbell</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#).into_bytes()
    }

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        let mut presentation = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            Vec::new(),
        );
        presentation.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".into(),
            "slides/slide1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/slides/slide1.xml").unwrap(),
            SLIDE_CONTENT_TYPE.into(),
            Vec::new(),
        )));
        package
    }

    fn value() -> ModernCommentPart {
        ModernCommentPart {
            slide_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rId9".into(),
            part_name: "/ppt/comments/modernComment1.xml".into(),
            comments: ModernCommentList::parse(&sdk_xml()).unwrap(),
        }
    }

    #[test]
    fn loads_microsoft_open_xml_sdk_documentation_specimen() {
        let parsed = ModernCommentList::parse(&sdk_xml()).unwrap();
        assert_eq!(parsed.comments.len(), 1);
        assert_eq!(parsed.comments[0].id, COMMENT);
        assert!(
            std::str::from_utf8(parsed.comments[0].text_body_xml.as_ref().unwrap())
                .unwrap()
                .contains("Needs more cowbell")
        );
        assert_eq!(
            ModernCommentList::parse(&parsed.to_xml().unwrap()).unwrap(),
            parsed
        );
    }

    #[test]
    fn package_round_trip_keeps_monikers_replies_and_extensions_inert() {
        let xml = format!(
            r#"<p188:cmLst xmlns:p188="{P188}" xmlns:pc="{PC}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:payload"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" status="resolved" created="2026-07-19T12:00:00+08:00" assignedTo="{AUTHOR}" complete="50%" title="Review"><pc:sldMkLst><pc:sldMk/></pc:sldMkLst><p188:pos x="10" y="-20"/><p188:replyLst><p188:reply id="{REPLY}" authorId="{AUTHOR}" created="2026-07-19T12:01:00+08:00"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p/></p188:txBody><p188:extLst><p:ext uri="{{A}}"><x:data relationship="rId999"/></p:ext></p188:extLst></p188:reply></p188:replyLst><p188:extLst><p:ext uri="{{B}}"><x:payload r:id="rId666" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></p:ext></p188:extLst></p188:cm></p188:cmLst>"#
        );
        let expected = ModernCommentList::parse(xml.as_bytes()).unwrap();
        let mut package = package();
        let mut part = value();
        part.comments = expected.clone();
        store_modern_comment(&mut package, &part).unwrap();
        let loaded = load_modern_comments(&package).unwrap();
        assert_eq!(loaded.len(), 1);
        assert_eq!(loaded[0].comments, expected);
        assert!(
            loaded[0].comments.comments[0]
                .extension_xml
                .as_ref()
                .unwrap()
                .windows(6)
                .any(|window| window == b"rId666")
        );
        assert!(
            package
                .get_part(&PackURI::new("/ppt/comments/modernComment1.xml").unwrap())
                .unwrap()
                .rels()
                .is_empty()
        );
    }

    #[test]
    fn rejects_hostile_or_schema_invalid_comment_xml() {
        let cases = [
            format!(r#"<!DOCTYPE x><p188:cmLst xmlns:p188="{P188}"/>"#),
            format!(r#"<x:cmLst xmlns:x="urn:wrong"/>"#),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="bad" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"/></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" status="pending" created="2024-12-30T20:26:06.503"/></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}" xmlns:pc="{PC}" xmlns:ac="{AC}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><pc:sldMkLst/><ac:deMkLst/></p188:cm></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><p188:txBody/><p188:replyLst/></p188:cm></p188:cmLst>"#
            ),
        ];
        for xml in cases {
            assert!(
                ModernCommentList::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        assert!(ModernCommentList::parse(&vec![b' '; MAX_BYTES + 1]).is_err());
    }

    #[test]
    fn rejects_invalid_package_graphs_and_failed_store_is_atomic() {
        let mut external = package();
        external
            .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                "https://invalid.example/comments.xml".into(),
                "rId9".into(),
                true,
            );
        assert!(load_modern_comments(&external).is_err());

        let mut wrong_source = package();
        wrong_source
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                "comments/modern.xml".into(),
                "rId9".into(),
                false,
            );
        assert!(load_modern_comments(&wrong_source).is_err());

        let mut orphan = package();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/comments/orphan.xml").unwrap(),
            MODERN_COMMENT_CONTENT_TYPE.into(),
            sdk_xml(),
        )));
        assert!(load_modern_comments(&orphan).is_err());

        let mut atomic = package();
        let mut invalid_value = value();
        invalid_value.comments.comments[0].id = "not-a-guid".into();
        let before_parts = atomic.iter_parts().count();
        let before_rels = atomic
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .rels()
            .len();
        assert!(store_modern_comment(&mut atomic, &invalid_value).is_err());
        assert_eq!(atomic.iter_parts().count(), before_parts);
        assert_eq!(
            atomic
                .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
                .unwrap()
                .rels()
                .len(),
            before_rels
        );
    }
}
