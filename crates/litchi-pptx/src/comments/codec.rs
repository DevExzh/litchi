//! Bounded XML parsing and writing for `PresentationML` legacy comment parts.

use super::{
    Author, Comment, Conformance, MAX_AUTHORS, MAX_COMMENTS_PER_SLIDE, MAX_DEPTH, MAX_NODES,
    MAX_PART_BYTES, MAX_STRING_BYTES, PML, STRICT_PML, invalid,
};
use crate::{Error, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

#[derive(Debug, Clone, PartialEq, Eq)]
struct Attribute {
    name: String,
    value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_comment_authors(xml: &[u8]) -> Result<Vec<Author>> {
    let root = parse_document(xml)?;
    let namespace = root_namespace(&root, "cmAuthorLst")?;
    no_attributes(&root, &[])?;
    whitespace_only(&root)?;
    if root.children.len() > MAX_AUTHORS {
        return Err(invalid("comment-author count exceeds limit"));
    }
    let mut authors = Vec::with_capacity(root.children.len());
    let mut ids = HashSet::new();
    for node in &root.children {
        require_name(node, namespace, "cmAuthor")?;
        whitespace_only(node)?;
        no_attributes(node, &["id", "name", "initials", "lastIdx", "clrIdx"])?;
        if node.children.len() > 1
            || node
                .children
                .first()
                .is_some_and(|child| child.namespace != namespace || child.name != "extLst")
        {
            return Err(invalid("cmAuthor permits only one trailing extLst"));
        }
        let author = Author {
            id: required_u32(node, "id")?,
            name: required(node, "name")?.to_owned(),
            initials: required(node, "initials")?.to_owned(),
            last_index: required_u32(node, "lastIdx")?,
            color_index: required_u32(node, "clrIdx")?,
        };
        if !ids.insert(author.id) {
            return Err(invalid(format!(
                "duplicate comment author ID {}",
                author.id
            )));
        }
        validate_string("comment author name", &author.name)?;
        validate_string("comment author initials", &author.initials)?;
        authors.push(author);
    }
    Ok(authors)
}

/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_slide_comments(xml: &[u8]) -> Result<Vec<Comment>> {
    let root = parse_document(xml)?;
    let namespace = root_namespace(&root, "cmLst")?;
    no_attributes(&root, &[])?;
    whitespace_only(&root)?;
    if root.children.len() > MAX_COMMENTS_PER_SLIDE {
        return Err(invalid("slide comment count exceeds limit"));
    }
    let mut comments = Vec::with_capacity(root.children.len());
    let mut keys = HashSet::new();
    for node in &root.children {
        require_name(node, namespace, "cm")?;
        whitespace_only(node)?;
        no_attributes(node, &["authorId", "dt", "idx"])?;
        if !(node.children.len() == 2 || node.children.len() == 3) {
            return Err(invalid(
                "cm requires pos and text, followed optionally by extLst",
            ));
        }
        let position = &node.children[0];
        require_name(position, namespace, "pos")?;
        no_attributes(position, &["x", "y"])?;
        require_empty(position)?;
        let text = &node.children[1];
        require_name(text, namespace, "text")?;
        no_attributes(text, &["xml:space"])?;
        if !text.children.is_empty() {
            return Err(invalid("comment text cannot contain elements"));
        }
        if let Some(extension_list) = node.children.get(2) {
            require_name(extension_list, namespace, "extLst")?;
        }
        let date_time = optional(node, "dt").map(ToOwned::to_owned);
        if let Some(value) = &date_time {
            validate_date_time(value)?;
        }
        let comment = Comment {
            author_id: required_u32(node, "authorId")?,
            date_time,
            index: required_u32(node, "idx")?,
            x: required_i64(position, "x")?,
            y: required_i64(position, "y")?,
            text: text.text.clone(),
        };
        if comment.index == 0 {
            return Err(invalid("comment index must be at least 1"));
        }
        validate_string("comment text", &comment.text)?;
        if !keys.insert((comment.author_id, comment.index)) {
            return Err(invalid("duplicate author/index comment key in slide"));
        }
        comments.push(comment);
    }
    Ok(comments)
}

/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write_comment_authors(authors: &[Author], conformance: Conformance) -> Result<Vec<u8>> {
    validate_authors(authors)?;
    let mut output = declaration();
    output.extend_from_slice(b"<p:cmAuthorLst xmlns:p=\"");
    escape(&mut output, conformance.namespace());
    output.extend_from_slice(b"\">");
    for author in authors {
        output.extend_from_slice(b"<p:cmAuthor");
        attribute(&mut output, "id", &author.id.to_string());
        attribute(&mut output, "name", &author.name);
        attribute(&mut output, "initials", &author.initials);
        attribute(&mut output, "lastIdx", &author.last_index.to_string());
        attribute(&mut output, "clrIdx", &author.color_index.to_string());
        output.extend_from_slice(b"/>");
    }
    output.extend_from_slice(b"</p:cmAuthorLst>");
    Ok(output)
}

/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write_slide_comments(comments: &[Comment], conformance: Conformance) -> Result<Vec<u8>> {
    validate_comment_list(comments)?;
    let mut output = declaration();
    output.extend_from_slice(b"<p:cmLst xmlns:p=\"");
    escape(&mut output, conformance.namespace());
    output.extend_from_slice(b"\">");
    for comment in comments {
        output.extend_from_slice(b"<p:cm");
        attribute(&mut output, "authorId", &comment.author_id.to_string());
        if let Some(date_time) = &comment.date_time {
            attribute(&mut output, "dt", date_time);
        }
        attribute(&mut output, "idx", &comment.index.to_string());
        output.extend_from_slice(b"><p:pos");
        attribute(&mut output, "x", &comment.x.to_string());
        attribute(&mut output, "y", &comment.y.to_string());
        output.extend_from_slice(b"/><p:text>");
        escape(&mut output, &comment.text);
        output.extend_from_slice(b"</p:text></p:cm>");
    }
    output.extend_from_slice(b"</p:cmLst>");
    Ok(output)
}

pub(super) fn validate_authors(authors: &[Author]) -> Result<()> {
    if authors.len() > MAX_AUTHORS {
        return Err(invalid("comment-author count exceeds limit"));
    }
    let mut ids = HashSet::new();
    for author in authors {
        validate_string("comment author name", &author.name)?;
        validate_string("comment author initials", &author.initials)?;
        if !ids.insert(author.id) {
            return Err(invalid(format!(
                "duplicate comment author ID {}",
                author.id
            )));
        }
    }
    Ok(())
}

pub(super) fn validate_comment_list(comments: &[Comment]) -> Result<()> {
    if comments.len() > MAX_COMMENTS_PER_SLIDE {
        return Err(invalid("slide comment count exceeds limit"));
    }
    let mut keys = HashSet::new();
    for comment in comments {
        if comment.index == 0 {
            return Err(invalid("comment index must be at least 1"));
        }
        validate_string("comment text", &comment.text)?;
        if let Some(date_time) = &comment.date_time {
            validate_date_time(date_time)?;
        }
        if !keys.insert((comment.author_id, comment.index)) {
            return Err(invalid("duplicate author/index comment key in slide"));
        }
    }
    Ok(())
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_PART_BYTES {
        return Err(invalid("presentation comment part is too large"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_PART_BYTES {
        return Err(invalid(
            "MCE-expanded presentation comment part is too large",
        ));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("presentation comment XML depth exceeds limit"));
                }
                let node = make_node(&element, namespace, decoder, &mut strings)?;
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(invalid("presentation comment XML node count exceeds limit"));
                }
                stack.push(node);
            },
            Event::Empty(element) => {
                let node = make_node(&element, namespace, decoder, &mut strings)?;
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(invalid("presentation comment XML node count exceeds limit"));
                }
                attach(node, &mut stack, &mut root)?;
            },
            Event::End(element) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
                if node.name.as_bytes() != element.local_name().as_ref() {
                    return Err(invalid("mismatched XML end element"));
                }
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                strings = strings
                    .checked_add(decoded.len())
                    .ok_or_else(|| invalid("XML string size overflow"))?;
                if strings > MAX_STRING_BYTES {
                    return Err(invalid("presentation comment XML strings exceed limit"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::CData(_) => {
                return Err(invalid("CDATA is rejected in presentation comment parts"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value =
                    if let Some(character) = reference.resolve_char_ref().map_err(xml_error)? {
                        character.to_string()
                    } else {
                        match name.as_ref() {
                            "amp" => "&".to_owned(),
                            "lt" => "<".to_owned(),
                            "gt" => ">".to_owned(),
                            "apos" => "'".to_owned(),
                            "quot" => "\"".to_owned(),
                            _ => {
                                return Err(invalid(format!(
                                    "custom XML entity '&{name};' is rejected"
                                )));
                            },
                        }
                    };
                strings = strings
                    .checked_add(value.len())
                    .ok_or_else(|| invalid("XML string size overflow"))?;
                if strings > MAX_STRING_BYTES {
                    return Err(invalid("presentation comment XML strings exceed limit"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity reference outside XML root"));
                }
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation comment XML"));
    }
    root.ok_or_else(|| invalid("missing presentation comment XML root"))
}

fn make_node(
    element: &BytesStart<'_>,
    namespace: ResolveResult<'_>,
    decoder: Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = match namespace {
        ResolveResult::Bound(value) => std::str::from_utf8(value.as_ref())
            .map_err(xml_error)?
            .to_owned(),
        ResolveResult::Unbound => String::new(),
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!(
                "unbound XML namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        },
    };
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute_value in element.attributes().with_checks(true) {
        let attribute_value = attribute_value.map_err(xml_error)?;
        let name = std::str::from_utf8(attribute_value.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate XML attribute"));
        }
        let value = attribute_value
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        *strings = strings
            .checked_add(name.len() + value.len())
            .ok_or_else(|| invalid("XML string size overflow"))?;
        if *strings > MAX_STRING_BYTES {
            return Err(invalid("presentation comment XML strings exceed limit"));
        }
        if name != "xmlns" && !name.starts_with("xmlns:") {
            attributes.push(Attribute { name, value });
        }
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn root_namespace<'a>(root: &'a Node, name: &str) -> Result<&'a str> {
    if root.name != name || !matches!(root.namespace.as_str(), PML | STRICT_PML) {
        return Err(invalid(format!("expected PresentationML {name} root")));
    }
    Ok(&root.namespace)
}

fn require_name(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected PresentationML {name}, got {}",
            node.name
        )))
    }
}

fn no_attributes(node: &Node, allowed: &[&str]) -> Result<()> {
    if let Some(attribute) = node
        .attributes
        .iter()
        .find(|attribute| !allowed.contains(&attribute.name.as_str()))
    {
        return Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )));
    }
    Ok(())
}

fn required<'a>(node: &'a Node, name: &str) -> Result<&'a str> {
    optional(node, name)
        .ok_or_else(|| invalid(format!("{} requires attribute '{name}'", node.name)))
}

fn optional<'a>(node: &'a Node, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}

fn required_u32(node: &Node, name: &str) -> Result<u32> {
    required(node, name)?.parse().map_err(|_err| {
        invalid(format!(
            "invalid unsigned integer '{name}' on {}",
            node.name
        ))
    })
}

fn required_i64(node: &Node, name: &str) -> Result<i64> {
    required(node, name)?
        .parse()
        .map_err(|_err| invalid(format!("invalid coordinate '{name}' on {}", node.name)))
}

fn whitespace_only(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}

fn require_empty(node: &Node) -> Result<()> {
    if node.children.is_empty() && node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{} must be empty", node.name)))
    }
}

fn validate_string(label: &str, value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(invalid(format!("{label} exceeds size limit")))
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

fn declaration() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec()
}

fn attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'"');
}

fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
