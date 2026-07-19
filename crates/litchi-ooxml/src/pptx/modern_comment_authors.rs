//! PowerPoint 2018 modern Comment Author parts.
//!
//! Author extensions are retained as XML and are never interpreted or used to
//! resolve relationships.

use super::modern_comments::{
    ModernCommentNamespaceDeclaration, ModernCommentPart, load_modern_comments,
};
use crate::common::mce::process_ooxml;
use crate::error::{OoxmlError, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

pub const MODERN_COMMENT_AUTHOR_CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.authors+xml";
pub const MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2018/10/relationships/authors";

const P188: &str = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 100_000;
const MAX_AUTHORS: usize = 65_536;
const MAX_STRING_BYTES: usize = 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModernCommentAuthor {
    pub id: String,
    pub name: String,
    pub initials: Option<String>,
    pub user_id: String,
    pub provider_id: String,
    pub namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
    /// Optional complete `p188:extLst` fragment retained inertly.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ModernCommentAuthorList {
    /// Prefix used for the 2018 PowerPoint namespace. Empty means default.
    pub root_prefix: String,
    pub namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
    pub authors: Vec<ModernCommentAuthor>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModernCommentAuthorPart {
    pub relationship_id: String,
    pub part_name: String,
    pub authors: ModernCommentAuthorList,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModernCommentGraph {
    pub authors: Option<ModernCommentAuthorPart>,
    pub comments: Vec<ModernCommentPart>,
}

impl ModernCommentAuthorList {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        parse_author_list(xml)
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        validate_author_model(self)?;
        let mut out = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
        open_tag(&mut out, &self.root_prefix, "authorLst");
        write_namespace_binding(&mut out, &self.root_prefix, P188);
        write_namespaces(&mut out, &self.namespace_declarations);
        if self.authors.is_empty() {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            for author in &self.authors {
                write_author(&mut out, &self.root_prefix, author);
            }
            close_tag(&mut out, &self.root_prefix, "authorLst");
        }
        if out.len() > MAX_BYTES {
            return Err(limit("serialized modern Comment Author bytes"));
        }
        parse_author_list(&out)?;
        Ok(out)
    }
}

pub fn load_modern_comment_authors(
    package: &OpcPackage,
) -> Result<Option<ModernCommentAuthorPart>> {
    let presentation = package.main_document_part()?;
    require_presentation_content_type(presentation.content_type())?;
    let presentation_name = presentation.partname().to_string();

    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "modern Comment Author relationship cannot originate at the package root",
        ));
    }
    for source in package.iter_parts() {
        if source.partname().as_str() != presentation_name.as_str()
            && source.rels().iter().any(|relationship| {
                relationship.reltype() == MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE
            })
        {
            return Err(invalid(
                "modern Comment Author relationship has non-Presentation source",
            ));
        }
    }

    let relationships: Vec<_> = presentation
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE)
        .collect();
    if relationships.len() > 1 {
        return Err(invalid(
            "Presentation has multiple modern Comment Author relationships",
        ));
    }
    let Some(relationship) = relationships.first().copied() else {
        if package
            .iter_parts()
            .any(|part| part.content_type() == MODERN_COMMENT_AUTHOR_CONTENT_TYPE)
        {
            return Err(invalid(
                "package contains an orphan modern Comment Author part",
            ));
        }
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(invalid(
            "modern Comment Author relationship cannot be external",
        ));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    require_content_type(part, MODERN_COMMENT_AUTHOR_CONTENT_TYPE)?;
    if !part.rels().is_empty() {
        return Err(invalid(
            "modern Comment Author part cannot have outbound relationships",
        ));
    }
    if package.iter_parts().any(|candidate| {
        candidate.content_type() == MODERN_COMMENT_AUTHOR_CONTENT_TYPE
            && candidate.partname() != &target
    }) {
        return Err(invalid(
            "package contains an orphan modern Comment Author part",
        ));
    }
    Ok(Some(ModernCommentAuthorPart {
        relationship_id: relationship.r_id().to_owned(),
        part_name: target.to_string(),
        authors: ModernCommentAuthorList::parse(part.blob())?,
    }))
}

pub fn load_modern_comment_graph(package: &OpcPackage) -> Result<ModernCommentGraph> {
    let authors = load_modern_comment_authors(package)?;
    let comments = load_modern_comments(package)?;
    validate_modern_comment_author_references(authors.as_ref(), &comments)?;
    Ok(ModernCommentGraph { authors, comments })
}

/// Validate modeled comment, reply, and assignment author references.
/// Author-looking values inside opaque extensions remain inert.
pub fn validate_modern_comment_author_references(
    authors: Option<&ModernCommentAuthorPart>,
    comments: &[ModernCommentPart],
) -> Result<()> {
    let has_references = comments.iter().any(|part| {
        part.comments.comments.iter().any(|comment| {
            true || !comment.replies.is_empty()
                || comment
                    .assigned_to
                    .as_ref()
                    .is_some_and(|ids| !ids.is_empty())
        })
    });
    if !has_references {
        return Ok(());
    }
    let authors = authors.ok_or_else(|| {
        invalid("modern comments reference authors but the package has no Author part")
    })?;
    let ids: HashSet<_> = authors
        .authors
        .authors
        .iter()
        .map(|author| author.id.as_str())
        .collect();
    for part in comments {
        for comment in &part.comments.comments {
            require_author_reference(&ids, &comment.author_id, "comment authorId")?;
            if let Some(assigned) = &comment.assigned_to {
                for author_id in assigned {
                    require_author_reference(&ids, author_id, "comment assignedTo")?;
                }
            }
            for reply in &comment.replies {
                require_author_reference(&ids, &reply.author_id, "reply authorId")?;
            }
        }
    }
    Ok(())
}

/// Add a new modern Comment Author part after validating the complete graph.
/// Existing Author parts are deliberately not overwritten.
pub fn store_modern_comment_authors(
    package: &mut OpcPackage,
    value: &ModernCommentAuthorPart,
) -> Result<()> {
    if load_modern_comment_authors(package)?.is_some() {
        return Err(invalid(
            "package already contains a modern Comment Author part",
        ));
    }
    let comments = load_modern_comments(package)?;
    validate_modern_comment_author_references(Some(value), &comments)?;
    validate_relationship_id(&value.relationship_id)?;
    let presentation = package.main_document_part()?;
    require_presentation_content_type(presentation.content_type())?;
    if presentation.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(
            "modern Comment Author relationship ID already exists",
        ));
    }
    let presentation_name = presentation.partname().clone();
    let part_name = PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    if package
        .iter_parts()
        .any(|part| part.partname() == &part_name)
    {
        return Err(invalid(format!("part '{part_name}' already exists")));
    }
    let xml = value.authors.to_xml()?;
    let target = part_name.relative_ref(presentation_name.base_uri());
    package.try_add_part(Box::new(BlobPart::new(
        part_name,
        MODERN_COMMENT_AUTHOR_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(&presentation_name)?
        .rels_mut()
        .add_relationship(
            MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
            target,
            value.relationship_id.clone(),
            false,
        );
    Ok(())
}

#[derive(Debug)]
enum FrameKind {
    Root,
    Author { index: usize, seen_extension: bool },
    RawExtension { index: usize, start: usize },
    Opaque,
}

#[derive(Debug)]
struct Frame {
    kind: FrameKind,
    namespace: String,
    local: String,
}

fn parse_author_list(xml: &[u8]) -> Result<ModernCommentAuthorList> {
    if xml.len() > MAX_BYTES {
        return Err(limit("modern Comment Author part bytes"));
    }
    let selected = process_ooxml(xml)?;
    if selected.len() > MAX_BYTES {
        return Err(limit("MCE-processed modern Comment Author bytes"));
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
    let mut authors = Vec::new();
    let mut nodes = 0usize;

    loop {
        let start = reader.buffer_position() as usize;
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
                    .ok_or_else(|| limit("modern Comment Author nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("modern Comment Author nodes"));
                }
                if stack.len() + 1 > MAX_DEPTH {
                    return Err(limit("modern Comment Author XML depth"));
                }
                let local = decode_name(element.local_name().as_ref())?;
                let kind = if stack.is_empty() {
                    if root_seen || root_closed || namespace != P188 || local != "authorLst" {
                        return Err(invalid("modern Comment Author root must be p188:authorLst"));
                    }
                    root_prefix = element_prefix(&element)?;
                    namespace_declarations =
                        namespace_declarations_from(&element, decoder, Some(&root_prefix))?;
                    no_non_namespace_attributes(&element)?;
                    root_seen = true;
                    FrameKind::Root
                } else {
                    child_frame(
                        &mut authors,
                        stack.last_mut().expect("nonempty stack"),
                        &element,
                        decoder,
                        &namespace,
                        &local,
                        start,
                    )?
                };
                let frame = Frame {
                    kind,
                    namespace,
                    local,
                };
                if empty {
                    attach_extension(
                        &frame.kind,
                        bytes,
                        reader.buffer_position() as usize,
                        &mut authors,
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
                    .ok_or_else(|| invalid("unexpected modern Comment Author closing element"))?;
                let local = decode_name(element.local_name().as_ref())?;
                if frame.namespace != namespace || frame.local != local {
                    return Err(invalid("mismatched modern Comment Author closing element"));
                }
                attach_extension(
                    &frame.kind,
                    bytes,
                    reader.buffer_position() as usize,
                    &mut authors,
                )?;
                if matches!(frame.kind, FrameKind::Root) {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                if !inside_extension(&stack) {
                    let decoded = text.decode().map_err(xml_error)?;
                    let value = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                    if !value.trim().is_empty() {
                        return Err(invalid("unexpected text in modern Comment Author metadata"));
                    }
                }
            },
            Event::CData(text) => {
                if !inside_extension(&stack) && !text.decode().map_err(xml_error)?.trim().is_empty()
                {
                    return Err(invalid(
                        "unexpected CDATA in modern Comment Author metadata",
                    ));
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
        return Err(invalid("unterminated modern Comment Author part"));
    }
    let value = ModernCommentAuthorList {
        root_prefix,
        namespace_declarations,
        authors,
    };
    validate_author_model(&value)?;
    Ok(value)
}

fn child_frame(
    authors: &mut Vec<ModernCommentAuthor>,
    parent: &mut Frame,
    element: &BytesStart<'_>,
    decoder: Decoder,
    namespace: &str,
    local: &str,
    start: usize,
) -> Result<FrameKind> {
    match &mut parent.kind {
        FrameKind::Root => {
            if namespace != P188 || local != "author" {
                return Err(invalid("authorLst permits only p188:author children"));
            }
            if authors.len() >= MAX_AUTHORS {
                return Err(limit("modern Comment Authors"));
            }
            let declarations = namespace_declarations_from(element, decoder, None)?;
            let attributes = known_attributes(
                element,
                decoder,
                &["id", "name", "initials", "userId", "providerId"],
            )?;
            authors.push(parse_author(attributes, declarations)?);
            Ok(FrameKind::Author {
                index: authors.len() - 1,
                seen_extension: false,
            })
        },
        FrameKind::Author {
            index,
            seen_extension,
        } => {
            if namespace != P188 || local != "extLst" || *seen_extension {
                return Err(invalid(
                    "modern Comment Author permits at most one p188:extLst child",
                ));
            }
            *seen_extension = true;
            validate_any_attributes(element, decoder)?;
            Ok(FrameKind::RawExtension {
                index: *index,
                start,
            })
        },
        FrameKind::RawExtension { .. } | FrameKind::Opaque => {
            validate_any_attributes(element, decoder)?;
            Ok(FrameKind::Opaque)
        },
    }
}

fn parse_author(
    attributes: HashMap<String, String>,
    namespace_declarations: Vec<ModernCommentNamespaceDeclaration>,
) -> Result<ModernCommentAuthor> {
    let id = required(&attributes, "id")?.to_owned();
    validate_guid(&id)?;
    Ok(ModernCommentAuthor {
        id,
        name: required(&attributes, "name")?.to_owned(),
        initials: attributes.get("initials").cloned(),
        user_id: required(&attributes, "userId")?.to_owned(),
        provider_id: required(&attributes, "providerId")?.to_owned(),
        namespace_declarations,
        extension_xml: None,
    })
}

fn attach_extension(
    kind: &FrameKind,
    bytes: &[u8],
    end: usize,
    authors: &mut [ModernCommentAuthor],
) -> Result<()> {
    let FrameKind::RawExtension { index, start } = kind else {
        return Ok(());
    };
    if *start > end || end > bytes.len() {
        return Err(invalid("invalid modern Comment Author extension bounds"));
    }
    authors[*index].extension_xml = Some(bytes[*start..end].to_vec());
    Ok(())
}

fn inside_extension(stack: &[Frame]) -> bool {
    stack
        .iter()
        .any(|frame| matches!(frame.kind, FrameKind::RawExtension { .. }))
}

fn validate_author_model(value: &ModernCommentAuthorList) -> Result<()> {
    validate_prefix(&value.root_prefix)?;
    validate_namespaces(&value.namespace_declarations, Some(&value.root_prefix))?;
    if value.authors.len() > MAX_AUTHORS {
        return Err(limit("modern Comment Authors"));
    }
    let mut ids = HashSet::new();
    for author in &value.authors {
        validate_guid(&author.id)?;
        if !ids.insert(author.id.as_str()) {
            return Err(invalid("duplicate modern Comment Author ID"));
        }
        bounded(&author.name)?;
        if let Some(initials) = &author.initials {
            bounded(initials)?;
        }
        bounded(&author.user_id)?;
        bounded(&author.provider_id)?;
        validate_namespaces(&author.namespace_declarations, None)?;
        if author
            .extension_xml
            .as_ref()
            .is_some_and(|xml| xml.len() > MAX_BYTES)
        {
            return Err(limit("modern Comment Author extension bytes"));
        }
    }
    Ok(())
}

fn write_author(out: &mut Vec<u8>, prefix: &str, author: &ModernCommentAuthor) {
    open_tag(out, prefix, "author");
    write_attr(out, "id", &author.id);
    write_attr(out, "name", &author.name);
    if let Some(initials) = &author.initials {
        write_attr(out, "initials", initials);
    }
    write_attr(out, "userId", &author.user_id);
    write_attr(out, "providerId", &author.provider_id);
    write_namespaces(out, &author.namespace_declarations);
    if let Some(extension) = &author.extension_xml {
        out.push(b'>');
        out.extend_from_slice(extension);
        close_tag(out, prefix, "author");
    } else {
        out.extend_from_slice(b"/>");
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
            return Err(invalid(format!("unexpected author attribute '{key}'")));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded(&value)?;
        if values.insert(key.clone(), value).is_some() {
            return Err(invalid(format!("duplicate author attribute '{key}'")));
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
            return Err(invalid(
                "unexpected attribute on modern Comment Author container",
            ));
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
            return Err(invalid(
                "modern Comment Author namespace prefix is declared twice",
            ));
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
        Err(invalid(format!(
            "invalid modern Comment Author GUID '{value}'"
        )))
    }
}

fn require_author_reference(ids: &HashSet<&str>, value: &str, label: &str) -> Result<()> {
    if ids.contains(value) {
        Ok(())
    } else {
        Err(invalid(format!(
            "{label} '{value}' does not resolve in the modern Comment Author part"
        )))
    }
}

fn validate_relationship_id(value: &str) -> Result<()> {
    bounded(value)?;
    if value.is_empty() || value.chars().any(char::is_whitespace) {
        Err(invalid(
            "modern Comment Author relationship ID must be nonempty without whitespace",
        ))
    } else {
        Ok(())
    }
}

fn required<'a>(attributes: &'a HashMap<String, String>, name: &str) -> Result<&'a str> {
    attributes.get(name).map(String::as_str).ok_or_else(|| {
        invalid(format!(
            "modern Comment Author is missing required '{name}'"
        ))
    })
}

fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("modern Comment Author string bytes"))
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

fn require_presentation_content_type(value: &str) -> Result<()> {
    if matches!(
        value,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
            | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "'{value}' is not a PresentationML main content type"
        )))
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

/// Find a modern comment author by its stable GUID without resolving its identity externally.
pub fn find_modern_comment_author(
    package: &OpcPackage,
    author_id: &str,
) -> Result<Option<ModernCommentAuthor>> {
    Ok(load_modern_comment_authors(package)?
        .and_then(|part| part.authors.authors.into_iter().find(|author| author.id == author_id)))
}

/// Add a modern comment author, allocating a collision-safe part and relationship if needed.
pub fn add_modern_comment_author(
    package: &mut OpcPackage,
    author: ModernCommentAuthor,
) -> Result<ModernCommentAuthorPart> {
    let mut graph = load_modern_comment_graph(package)?;
    if graph
        .authors
        .as_ref()
        .is_some_and(|part| part.authors.authors.iter().any(|item| item.id == author.id))
    {
        return Err(invalid(format!("duplicate modern comment author ID {}", author.id)));
    }

    if let Some(mut part) = graph.authors.take() {
        part.authors.authors.push(author);
        commit_modern_comment_authors(package, &part, &graph.comments)?;
        return Ok(part);
    }

    let presentation_name = package.main_document_part()?.partname().clone();
    let part_name = next_modern_author_part_name(package)?;
    let relationship_id = next_modern_author_relationship_id(package, &presentation_name)?;
    let part = ModernCommentAuthorPart {
        relationship_id: relationship_id.clone(),
        part_name: part_name.to_string(),
        authors: ModernCommentAuthorList {
            root_prefix: "p188".into(),
            namespace_declarations: Vec::new(),
            authors: vec![author],
        },
    };
    validate_modern_comment_author_references(Some(&part), &graph.comments)?;
    let xml = part.authors.to_xml()?;
    ModernCommentAuthorList::parse(&xml)?;
    package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
        part_name.clone(),
        MODERN_COMMENT_AUTHOR_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(&presentation_name)?
        .rels_mut()
        .add_relationship(
            MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
            part_name.relative_ref(presentation_name.base_uri()),
            relationship_id,
            false,
        );
    let _ = package.clear_digital_signatures();
    Ok(part)
}

/// Update a modern comment author while keeping its stable GUID unchanged.
pub fn update_modern_comment_author<F>(
    package: &mut OpcPackage,
    author_id: &str,
    update: F,
) -> Result<bool>
where
    F: FnOnce(&mut ModernCommentAuthor),
{
    let graph = load_modern_comment_graph(package)?;
    let Some(mut part) = graph.authors.clone() else { return Ok(false); };
    let Some(author) = part.authors.authors.iter_mut().find(|item| item.id == author_id) else {
        return Ok(false);
    };
    update(author);
    if author.id != author_id {
        return Err(invalid("modern comment author update cannot change its ID"));
    }
    commit_modern_comment_authors(package, &part, &graph.comments)?;
    Ok(true)
}

/// Replace a modern comment author while keeping its stable GUID unchanged.
pub fn replace_modern_comment_author(
    package: &mut OpcPackage,
    author_id: &str,
    replacement: ModernCommentAuthor,
) -> Result<bool> {
    if replacement.id != author_id {
        return Err(invalid("replacement modern comment author ID must match"));
    }
    update_modern_comment_author(package, author_id, move |author| *author = replacement)
}

/// Remove an unreferenced modern comment author.
pub fn remove_modern_comment_author(package: &mut OpcPackage, author_id: &str) -> Result<bool> {
    let graph = load_modern_comment_graph(package)?;
    let Some(mut part) = graph.authors.clone() else { return Ok(false); };
    let Some(index) = part.authors.authors.iter().position(|author| author.id == author_id) else {
        return Ok(false);
    };
    if modern_author_is_referenced(&graph.comments, author_id) {
        return Err(invalid(format!("modern comment author {author_id} is still referenced")));
    }
    part.authors.authors.remove(index);
    if !part.authors.authors.is_empty() {
        commit_modern_comment_authors(package, &part, &graph.comments)?;
        return Ok(true);
    }

    let presentation_name = package.main_document_part()?.partname().clone();
    let part_name = PackURI::new(&part.part_name).map_err(invalid)?;
    package
        .get_part_mut(&presentation_name)?
        .rels_mut()
        .remove(&part.relationship_id);
    if !modern_author_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

/// Reorder modern comment authors by a complete, duplicate-free GUID list.
pub fn reorder_modern_comment_authors(
    package: &mut OpcPackage,
    ordered_author_ids: &[String],
) -> Result<Vec<ModernCommentAuthor>> {
    let graph = load_modern_comment_graph(package)?;
    let Some(mut part) = graph.authors.clone() else {
        if ordered_author_ids.is_empty() { return Ok(Vec::new()); }
        return Err(invalid("modern comment author part is missing"));
    };
    if ordered_author_ids.len() != part.authors.authors.len() {
        return Err(invalid("modern author reorder must contain every author"));
    }
    let mut remaining = std::collections::HashMap::new();
    for author in part.authors.authors.drain(..) {
        if remaining.insert(author.id.clone(), author).is_some() {
            return Err(invalid("duplicate modern comment author ID"));
        }
    }
    let mut ordered = Vec::with_capacity(ordered_author_ids.len());
    for id in ordered_author_ids {
        let author = remaining
            .remove(id)
            .ok_or_else(|| invalid(format!("unknown or duplicate modern comment author ID {id}")))?;
        ordered.push(author);
    }
    if !remaining.is_empty() {
        return Err(invalid("modern author reorder must contain every author"));
    }
    part.authors.authors = ordered.clone();
    commit_modern_comment_authors(package, &part, &graph.comments)?;
    Ok(ordered)
}

fn commit_modern_comment_authors(
    package: &mut OpcPackage,
    part: &ModernCommentAuthorPart,
    comments: &[ModernCommentPart],
) -> Result<()> {
    validate_modern_comment_author_references(Some(part), comments)?;
    let xml = part.authors.to_xml()?;
    ModernCommentAuthorList::parse(&xml)?;
    let part_name = PackURI::new(&part.part_name).map_err(invalid)?;
    package.get_part_mut(&part_name)?.set_blob(xml);
    let _ = package.clear_digital_signatures();
    Ok(())
}

fn modern_author_is_referenced(comments: &[ModernCommentPart], author_id: &str) -> bool {
    comments.iter().any(|part| part.comments.comments.iter().any(|comment| {
        comment.author_id == author_id
            || comment.assigned_to.as_ref().is_some_and(|ids| ids.iter().any(|id| id == author_id))
            || comment.replies.iter().any(|reply| reply.author_id == author_id)
    }))
}

fn next_modern_author_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 1..=65_537u32 {
        let candidate = PackURI::new(&format!("/ppt/authors/author{suffix}.xml")).map_err(invalid)?;
        if package.get_part(&candidate).is_err() { return Ok(candidate); }
    }
    Err(invalid("no free modern comment author part name"))
}

fn next_modern_author_relationship_id(
    package: &OpcPackage,
    presentation_name: &PackURI,
) -> Result<String> {
    let relationships = package.get_part(presentation_name)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdModernAuthors{suffix}");
        if relationships.get(&candidate).is_none() { return Ok(candidate); }
    }
    Err(invalid("no free modern comment author relationship ID"))
}

fn modern_author_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
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
    use super::super::modern_comments::{ModernCommentList, store_modern_comment};
    use super::*;

    const AUTHOR: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
    const OTHER: &str = "{0B2043D4-0908-4C42-8A79-51EA2CC309F7}";
    const COMMENT: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";

    fn sdk_author_xml() -> Vec<u8> {
        format!(r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="Ada Lovelace" initials="AL" userId="ada@example.com::4b640067-2830-4c10-9c4f-5879bb2e41d1" providerId=""/></p188:authorLst>"#).into_bytes()
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
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            Vec::new(),
        )));
        package
    }

    fn author_part() -> ModernCommentAuthorPart {
        ModernCommentAuthorPart {
            relationship_id: "rId8".into(),
            part_name: "/ppt/authors/author1.xml".into(),
            authors: ModernCommentAuthorList::parse(&sdk_author_xml()).unwrap(),
        }
    }

    fn comment_part(author: &str) -> ModernCommentPart {
        let xml = format!(
            r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{author}" created="2024-12-30T20:26:06.503" assignedTo="{author}"/></p188:cmLst>"#
        );
        ModernCommentPart {
            slide_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rId9".into(),
            part_name: "/ppt/comments/modernComment1.xml".into(),
            comments: ModernCommentList::parse(xml.as_bytes()).unwrap(),
        }
    }

    #[test]
    fn loads_open_xml_sdk_shaped_author_specimen() {
        let parsed = ModernCommentAuthorList::parse(&sdk_author_xml()).unwrap();
        assert_eq!(parsed.authors.len(), 1);
        assert_eq!(parsed.authors[0].name, "Ada Lovelace");
        assert_eq!(parsed.authors[0].provider_id, "");
        assert_eq!(
            ModernCommentAuthorList::parse(&parsed.to_xml().unwrap()).unwrap(),
            parsed
        );
    }

    #[test]
    fn author_and_comment_package_graph_round_trip_and_resolve() {
        let extension = format!(r#"<p188:extLst xmlns:p188="{P188}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:payload"><p:ext uri="{{A}}"><x:data authorId="{OTHER}" relationship="rId999"/></p:ext></p188:extLst>"#).into_bytes();
        let mut authors = author_part();
        authors.authors.authors[0].extension_xml = Some(extension.clone());
        let mut package = package();
        store_modern_comment_authors(&mut package, &authors).unwrap();
        store_modern_comment(&mut package, &comment_part(AUTHOR)).unwrap();
        let graph = load_modern_comment_graph(&package).unwrap();
        assert_eq!(
            graph.authors.unwrap().authors.authors[0].extension_xml,
            Some(extension)
        );
        assert_eq!(graph.comments.len(), 1);
    }

    #[test]
    fn rejects_hostile_author_grammar_and_unresolved_modeled_references() {
        let cases = [
            format!(r#"<!DOCTYPE x><p188:authorLst xmlns:p188="{P188}"/>"#),
            "<x:authorLst xmlns:x=\"urn:wrong\"/>".into(),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="bad" name="A" userId="u" providerId="p"/></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u"/></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u" providerId="p"><p188:extLst/><p188:extLst/></p188:author></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u" providerId="p"/><p188:author id="{AUTHOR}" name="B" userId="v" providerId="p"/></p188:authorLst>"#
            ),
        ];
        for xml in cases {
            assert!(
                ModernCommentAuthorList::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        assert!(ModernCommentAuthorList::parse(&vec![b' '; MAX_BYTES + 1]).is_err());

        let authors = author_part();
        assert!(
            validate_modern_comment_author_references(Some(&authors), &[comment_part(OTHER)])
                .is_err()
        );
        assert!(validate_modern_comment_author_references(None, &[comment_part(AUTHOR)]).is_err());
    }

    #[test]
    fn rejects_author_package_graphs_and_failed_store_is_atomic() {
        let mut external = package();
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
                "https://invalid.example/authors.xml".into(),
                "rId8".into(),
                true,
            );
        assert!(load_modern_comment_authors(&external).is_err());

        let mut orphan = package();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/authors/orphan.xml").unwrap(),
            MODERN_COMMENT_AUTHOR_CONTENT_TYPE.into(),
            sdk_author_xml(),
        )));
        assert!(load_modern_comment_authors(&orphan).is_err());

        let mut outbound = package();
        store_modern_comment_authors(&mut outbound, &author_part()).unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/authors/author1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "other.xml".into(),
                "rId1".into(),
                false,
            );
        assert!(load_modern_comment_authors(&outbound).is_err());

        let mut atomic = package();
        store_modern_comment(&mut atomic, &comment_part(OTHER)).unwrap();
        let before_parts = atomic.iter_parts().count();
        let before_rels = atomic
            .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels()
            .len();
        assert!(store_modern_comment_authors(&mut atomic, &author_part()).is_err());
        assert_eq!(atomic.iter_parts().count(), before_parts);
        assert_eq!(
            atomic
                .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
                .unwrap()
                .rels()
                .len(),
            before_rels
        );
    }
}
