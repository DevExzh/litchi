//! Bounded XML parsing and writing for `PresentationML` legacy comment parts.

use super::{
    Author, Comment, Conformance, Limits, MAX_AUTHORS, MAX_COMMENTS_PER_SLIDE, PML, STRICT_PML,
    invalid, limit,
};
use crate::leniency::{CommentDefect, MarkupDefect, ToleranceReport};
use crate::{Error, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, process_markup_compatibility};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::{borrow::Cow, collections::HashSet};

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
pub fn parse_comment_authors(xml: &[u8], limits: Limits) -> Result<Vec<Author>> {
    parse_comment_authors_with(xml, limits, None)
}

/// [`parse_comment_authors`], recording skipped processing instructions in
/// `tolerance` when one is given.
pub(super) fn parse_comment_authors_with(
    xml: &[u8],
    limits: Limits,
    tolerance: Option<&mut ToleranceReport>,
) -> Result<Vec<Author>> {
    let (root, instructions) = parse_document(xml, limits)?;
    if instructions && let Some(tolerance) = tolerance {
        tolerance.record(MarkupDefect::ProcessingInstruction, 0);
    }
    let namespace = root_namespace(&root, "cmAuthorLst")?;
    no_attributes(&root, &[])?;
    whitespace_only(&root)?;
    let maximum = MAX_AUTHORS.min(limits.max_nodes);
    if root.children.len() > maximum {
        return Err(limit(
            crate::error::CodecLimitKind::Authors,
            root.children.len(),
            maximum,
        ));
    }
    let mut authors = Vec::new();
    authors
        .try_reserve_exact(root.children.len())
        .map_err(|source| Error::Allocation {
            resource: "comment author records",
            source,
        })?;
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
        validate_string_with_limit("comment author name", &author.name, limits.max_string_bytes)?;
        validate_string_with_limit(
            "comment author initials",
            &author.initials,
            limits.max_string_bytes,
        )?;
        authors.push(author);
    }
    Ok(authors)
}

/// Parse a legacy slide comment list (`p:cmLst`), applying the
/// `PowerPoint`-compatible repairs of audit finding 6.12: a comment with
/// `idx="0"` is kept as written, and a comment repeating an earlier
/// `(authorId, idx)` key is kept after it (see [`CommentDefect`]). Use
/// [`parse_slide_comments_with_tolerance`] to collect what was repaired, or
/// [`parse_slide_comments_strict`] to refuse such a list.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_slide_comments(xml: &[u8], limits: Limits) -> Result<Vec<Comment>> {
    parse_slide_comments_with_tolerance(xml, limits, &mut ToleranceReport::default())
}

/// [`parse_slide_comments`], appending each [`CommentDefect`] repair to
/// `tolerance` (no part is stamped; the caller knows the part name).
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_slide_comments_with_tolerance(
    xml: &[u8],
    limits: Limits,
    tolerance: &mut ToleranceReport,
) -> Result<Vec<Comment>> {
    parse_slide_comments_inner(xml, limits, Some(tolerance))
}

/// Parse a legacy slide comment list without any repair: a comment index of
/// `0` or a repeated `(authorId, idx)` key is an error. The comment edit
/// functions load through this contract, so they never rewrite a list whose
/// read needed a repair.
///
/// # Errors
///
/// Returns an error if the input cannot be read, is malformed, or would
/// need a [`CommentDefect`] repair.
pub fn parse_slide_comments_strict(xml: &[u8], limits: Limits) -> Result<Vec<Comment>> {
    parse_slide_comments_inner(xml, limits, None)
}

fn parse_slide_comments_inner(
    xml: &[u8],
    limits: Limits,
    mut tolerance: Option<&mut ToleranceReport>,
) -> Result<Vec<Comment>> {
    let (root, instructions) = parse_document(xml, limits)?;
    if instructions && let Some(tolerance) = tolerance.as_deref_mut() {
        tolerance.record(MarkupDefect::ProcessingInstruction, 0);
    }
    let namespace = root_namespace(&root, "cmLst")?;
    no_attributes(&root, &[])?;
    whitespace_only(&root)?;
    let maximum = MAX_COMMENTS_PER_SLIDE.min(limits.max_nodes);
    if root.children.len() > maximum {
        return Err(limit(
            crate::error::CodecLimitKind::CommentEntries,
            root.children.len(),
            maximum,
        ));
    }
    let mut comments = Vec::new();
    comments
        .try_reserve_exact(root.children.len())
        .map_err(|source| Error::Allocation {
            resource: "slide comment records",
            source,
        })?;
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
        validate_string_with_limit("comment text", &comment.text, limits.max_string_bytes)?;
        if comment.index == 0 {
            let Some(tolerance) = tolerance.as_deref_mut() else {
                return Err(invalid("comment index must be at least 1"));
            };
            tolerance.record_comment(CommentDefect::ZeroIndex, comment.author_id, 0);
        }
        if !keys.insert((comment.author_id, comment.index)) {
            let Some(tolerance) = tolerance.as_deref_mut() else {
                return Err(invalid("duplicate author/index comment key in slide"));
            };
            tolerance.record_comment(
                CommentDefect::DuplicateKey,
                comment.author_id,
                comment.index,
            );
        }
        comments.push(comment);
    }
    Ok(comments)
}

/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write_comment_authors(
    authors: &[Author],
    conformance: Conformance,
    limits: Limits,
) -> Result<Vec<u8>> {
    validate_authors_with_limit(authors, limits.max_string_bytes, limits.max_nodes)?;
    let mut output = BoundedOutput::new(limits.max_output_bytes);
    output.write(declaration())?;
    output.write(b"<p:cmAuthorLst xmlns:p=\"")?;
    output.escape(conformance.namespace())?;
    output.write(b"\">")?;
    for author in authors {
        output.write(b"<p:cmAuthor")?;
        output.attribute("id", &author.id.to_string())?;
        output.attribute("name", &author.name)?;
        output.attribute("initials", &author.initials)?;
        output.attribute("lastIdx", &author.last_index.to_string())?;
        output.attribute("clrIdx", &author.color_index.to_string())?;
        output.write(b"/>")?;
    }
    output.write(b"</p:cmAuthorLst>")?;
    output.finish()
}

/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write_slide_comments(
    comments: &[Comment],
    conformance: Conformance,
    limits: Limits,
) -> Result<Vec<u8>> {
    validate_comment_list_with_limit(comments, limits.max_string_bytes, limits.max_nodes)?;
    let mut output = BoundedOutput::new(limits.max_output_bytes);
    output.write(declaration())?;
    output.write(b"<p:cmLst xmlns:p=\"")?;
    output.escape(conformance.namespace())?;
    output.write(b"\">")?;
    for comment in comments {
        output.write(b"<p:cm")?;
        output.attribute("authorId", &comment.author_id.to_string())?;
        if let Some(date_time) = &comment.date_time {
            output.attribute("dt", date_time)?;
        }
        output.attribute("idx", &comment.index.to_string())?;
        output.write(b"><p:pos")?;
        output.attribute("x", &comment.x.to_string())?;
        output.attribute("y", &comment.y.to_string())?;
        output.write(b"/><p:text>")?;
        output.escape(&comment.text)?;
        output.write(b"</p:text></p:cm>")?;
    }
    output.write(b"</p:cmLst>")?;
    output.finish()
}

pub(super) fn validate_authors_with_limit(
    authors: &[Author],
    max_string_bytes: usize,
    max_nodes: usize,
) -> Result<()> {
    let maximum = MAX_AUTHORS.min(max_nodes);
    if authors.len() > maximum {
        return Err(limit(
            crate::error::CodecLimitKind::Authors,
            authors.len(),
            maximum,
        ));
    }
    let mut ids = HashSet::new();
    for author in authors {
        validate_string_with_limit("comment author name", &author.name, max_string_bytes)?;
        validate_string_with_limit(
            "comment author initials",
            &author.initials,
            max_string_bytes,
        )?;
        if !ids.insert(author.id) {
            return Err(invalid(format!(
                "duplicate comment author ID {}",
                author.id
            )));
        }
    }
    Ok(())
}

pub(super) fn validate_comment_list_with_limit(
    comments: &[Comment],
    max_string_bytes: usize,
    max_nodes: usize,
) -> Result<()> {
    let maximum = MAX_COMMENTS_PER_SLIDE.min(max_nodes);
    if comments.len() > maximum {
        return Err(limit(
            crate::error::CodecLimitKind::CommentEntries,
            comments.len(),
            maximum,
        ));
    }
    let mut keys = HashSet::new();
    for comment in comments {
        if comment.index == 0 {
            return Err(invalid("comment index must be at least 1"));
        }
        validate_string_with_limit("comment text", &comment.text, max_string_bytes)?;
        if let Some(date_time) = &comment.date_time {
            validate_date_time(date_time)?;
        }
        if !keys.insert((comment.author_id, comment.index)) {
            return Err(invalid("duplicate author/index comment key in slide"));
        }
    }
    Ok(())
}

/// Parse a comment part into its node tree, and whether processing
/// instructions were skipped.
fn parse_document(xml: &[u8], limits: Limits) -> Result<(Node, bool)> {
    if xml.len() > limits.max_input_bytes {
        return Err(Error::Limit {
            kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::XmlBytes),
            observed: xml.len(),
            maximum: limits.max_input_bytes,
        });
    }
    let (processed, instructions) = process_comment_mce(xml, limits)?;
    if processed.len() > limits.max_input_bytes {
        return Err(Error::Limit {
            kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::XmlBytes),
            observed: processed.len(),
            maximum: limits.max_input_bytes,
        });
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
                if stack.len() >= limits.max_depth {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        stack.len().checked_add(1).unwrap_or(usize::MAX),
                        limits.max_depth,
                    ));
                }
                let node = make_node(
                    &element,
                    namespace,
                    decoder,
                    &mut strings,
                    limits.max_string_bytes,
                )?;
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes,
                    )
                })?;
                if nodes > limits.max_nodes {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes,
                    ));
                }
                stack.push(node);
            },
            Event::Empty(element) => {
                let node = make_node(
                    &element,
                    namespace,
                    decoder,
                    &mut strings,
                    limits.max_string_bytes,
                )?;
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes,
                    )
                })?;
                if nodes > limits.max_nodes {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes,
                    ));
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
                strings = strings.checked_add(decoded.len()).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::StringBytes,
                        usize::MAX,
                        limits.max_string_bytes,
                    )
                })?;
                if strings > limits.max_string_bytes {
                    return Err(limit(
                        crate::error::CodecLimitKind::StringBytes,
                        strings,
                        limits.max_string_bytes,
                    ));
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
            Event::DocType(_) => {
                return Err(invalid("DTD is rejected in presentation comment parts"));
            },
            // Processing instructions are inert markup.  Keep accepting them
            // so source-bound collaboration edits can retain their exact
            // surrounding bytes.
            Event::PI(_) => {},
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
                strings = strings.checked_add(value.len()).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::StringBytes,
                        usize::MAX,
                        limits.max_string_bytes,
                    )
                })?;
                if strings > limits.max_string_bytes {
                    return Err(limit(
                        crate::error::CodecLimitKind::StringBytes,
                        strings,
                        limits.max_string_bytes,
                    ));
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
        .map(|root| (root, instructions))
}

/// The common MCE processor rejects processing instructions while the legacy
/// comments owner treats them as inert source markup. Remove PIs only from
/// the temporary semantic input; source-bound collaboration codecs continue
/// to retain the original bytes for exact no-ops and inverse patches. The
/// flag reports whether any were removed (audit finding 6.3).
fn process_comment_mce(xml: &[u8], limits: Limits) -> Result<(Cow<'_, [u8]>, bool)> {
    let (filtered, found) = crate::markup::without_processing_instructions(xml)?;
    let mce_limits = MceLimits {
        max_input_bytes: limits.max_input_bytes,
        max_output_bytes: limits.max_input_bytes,
        max_depth: limits.max_depth,
        ..MceLimits::default()
    };
    let processed = match filtered {
        Cow::Borrowed(xml) => {
            process_markup_compatibility(xml, &Capabilities::ooxml_baseline(), &mce_limits)?.xml
        },
        Cow::Owned(filtered) => Cow::Owned(
            process_markup_compatibility(&filtered, &Capabilities::ooxml_baseline(), &mce_limits)?
                .xml
                .into_owned(),
        ),
    };
    Ok((processed, found))
}

fn make_node(
    element: &BytesStart<'_>,
    namespace: ResolveResult<'_>,
    decoder: Decoder,
    strings: &mut usize,
    max_string_bytes: usize,
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
    for attribute_value in element.checked_attributes() {
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
        let attribute_bytes = name.len().checked_add(value.len()).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::StringBytes,
                usize::MAX,
                max_string_bytes,
            )
        })?;
        *strings = strings.checked_add(attribute_bytes).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::StringBytes,
                usize::MAX,
                max_string_bytes,
            )
        })?;
        if *strings > max_string_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::StringBytes,
                *strings,
                max_string_bytes,
            ));
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

fn validate_string_with_limit(_label: &str, value: &str, max_string_bytes: usize) -> Result<()> {
    if value.len() <= max_string_bytes {
        Ok(())
    } else {
        Err(Error::Limit {
            kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::StringBytes),
            observed: value.len(),
            maximum: max_string_bytes,
        })
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

fn declaration() -> &'static [u8] {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
}

struct BoundedOutput {
    bytes: Vec<u8>,
    limit: usize,
}

impl BoundedOutput {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn write(&mut self, value: &[u8]) -> Result<()> {
        let next = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or_else(|| output_limit(usize::MAX, self.limit))?;
        if next > self.limit {
            return Err(output_limit(next, self.limit));
        }
        let spare = self.bytes.capacity().saturating_sub(self.bytes.len());
        if value.len() > spare {
            self.bytes
                .try_reserve_exact(value.len())
                .map_err(|source| Error::Allocation {
                    resource: "serialized presentation comment XML",
                    source,
                })?;
        }
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn escape(&mut self, value: &str) -> Result<()> {
        for character in value.chars() {
            match character {
                '&' => self.write(b"&amp;")?,
                '<' => self.write(b"&lt;")?,
                '>' => self.write(b"&gt;")?,
                '"' => self.write(b"&quot;")?,
                '\t' => self.write(b"&#x9;")?,
                '\n' => self.write(b"&#xA;")?,
                '\r' => self.write(b"&#xD;")?,
                _ => {
                    let mut bytes = [0; 4];
                    self.write(character.encode_utf8(&mut bytes).as_bytes())?;
                },
            }
        }
        Ok(())
    }

    fn attribute(&mut self, name: &str, value: &str) -> Result<()> {
        self.write(b" ")?;
        self.write(name.as_bytes())?;
        self.write(b"=\"")?;
        self.escape(value)?;
        self.write(b"\"")
    }

    fn finish(self) -> Result<Vec<u8>> {
        Ok(self.bytes)
    }
}

fn output_limit(observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::XmlBytes),
        observed,
        maximum,
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
