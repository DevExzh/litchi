//! Inert PresentationML programmable tags and slide relationship discovery.
//!
//! Names are selected semantically and values are always inert strings. The
//! module never interprets a value as XML, a path, a command, or a relationship
//! target.
//!
//! ```
//! use litchi_pptx::tag::{List, Tag};
//!
//! let mut tags = List::new();
//! tags.add(Tag::new("Owner", "Alice")?)?;
//! assert_eq!(tags.get("owner")?.value(), "Alice");
//! assert!(tags.get(1_usize).is_err());
//! # Ok::<(), litchi_pptx::Error>(())
//! ```

use crate::{Error, Result};
use caseless::Caseless;
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::{OpcPackage, PackURI, Part as OpcPart};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use unicode_normalization::UnicodeNormalization;

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const PML_TEXT: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_TEXT: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const TAG_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
const STRICT_TAG_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/tags";

/// Content type of a PresentationML programmable-tag part.
pub const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";

const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_TAGS: usize = 16_384;
const MAX_TAG_PARTS: usize = 1_024;
const XML_DECL: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const ROOT_OPEN: &[u8] = b"<p:tagLst xmlns:p=\"";
const ROOT_EMPTY_CLOSE: &[u8] = b"/>";
const ROOT_CHILDREN_OPEN: &[u8] = b">";
const ROOT_CLOSE: &[u8] = b"</p:tagLst>";
const TAG_OPEN: &[u8] = b"<p:tag";
const TAG_CLOSE: &[u8] = b"/>";
const MAX_NAMESPACE_BYTES: usize = if PML_TEXT.len() > STRICT_TEXT.len() {
    PML_TEXT.len()
} else {
    STRICT_TEXT.len()
};
const ROOT_PREFIX_BYTES: usize = XML_DECL.len() + ROOT_OPEN.len() + MAX_NAMESPACE_BYTES + 1;
const EMPTY_WIRE_BYTES: usize = ROOT_PREFIX_BYTES + ROOT_EMPTY_CLOSE.len();
const NONEMPTY_WIRE_BYTES: usize = ROOT_PREFIX_BYTES + ROOT_CHILDREN_OPEN.len() + ROOT_CLOSE.len();

/// Namespace profile used when a detached list is serialized.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    /// ECMA-376 transitional namespace profile.
    Transitional,
    /// ISO/IEC 29500 strict namespace profile.
    Strict,
}

impl Conformance {
    fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => PML_TEXT,
            Self::Strict => STRICT_TEXT,
        }
    }
}

/// Lossless low-level markup retained by the bounded codec.
pub mod raw {
    use super::{Result, bounded_text, validate_qname};

    /// One inert extension or namespace attribute.
    #[derive(Debug, Clone, PartialEq, Eq)]
    pub struct Attr {
        pub(super) qualified_name: String,
        pub(super) value: String,
    }

    impl Attr {
        /// Construct a bounded XML attribute without interpreting its value.
        pub fn new(qualified_name: impl Into<String>, value: impl Into<String>) -> Result<Self> {
            let qualified_name = qualified_name.into();
            let value = value.into();
            validate_qname(&qualified_name)?;
            bounded_text(&value, "tag attribute")?;
            Ok(Self {
                qualified_name,
                value,
            })
        }

        /// Return the retained qualified spelling.
        pub fn qualified_name(&self) -> &str {
            &self.qualified_name
        }

        /// Return the inert retained value.
        pub fn value(&self) -> &str {
            &self.value
        }
    }
}

/// A stable checked selector. Semantic names are the primary entry point;
/// numeric positions remain available for source-order workflows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Key<'a> {
    /// Unicode-caseless tag name.
    Name(&'a str),
    /// Zero-based source-order position.
    Index(usize),
}

impl<'a> From<&'a str> for Key<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<usize> for Key<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

/// One inert programmable tag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Tag {
    name: String,
    value: String,
    namespaces: Vec<raw::Attr>,
    attrs: Vec<raw::Attr>,
    wire_len: usize,
}

impl Tag {
    /// Construct an owned name/value pair.
    pub fn new(name: impl Into<String>, value: impl Into<String>) -> Result<Self> {
        let name = name.into();
        let value = value.into();
        bounded_text(&name, "tag name")?;
        bounded_text(&value, "tag value")?;
        let wire_len = tag_wire_len_parts(&name, &value, &[], &[])?;
        ensure_tag_budget(wire_len)?;
        Ok(Self {
            name,
            value,
            namespaces: Vec::new(),
            attrs: Vec::new(),
            wire_len,
        })
    }

    /// Return the producer spelling of the name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the inert string value.
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Return retained extension namespace declarations.
    pub fn namespaces(&self) -> &[raw::Attr] {
        &self.namespaces
    }

    /// Return retained extension attributes.
    pub fn attrs(&self) -> &[raw::Attr] {
        &self.attrs
    }

    /// Replace the standalone tag's name after bounded validation.
    ///
    /// In-list renames should use [`List::replace`] so uniqueness is checked
    /// against the complete list.
    pub fn rename(&mut self, name: impl Into<String>) -> Result<()> {
        let name = name.into();
        bounded_text(&name, "tag name")?;
        let wire_len =
            replace_wire_component(self.wire_len, escaped_len(&self.name)?, escaped_len(&name)?)?;
        ensure_tag_budget(wire_len)?;
        self.name = name;
        self.wire_len = wire_len;
        Ok(())
    }

    /// Replace the inert value, returning the previous allocation.
    pub fn set_value(&mut self, value: impl Into<String>) -> Result<String> {
        let value = value.into();
        bounded_text(&value, "tag value")?;
        let wire_len = replace_wire_component(
            self.wire_len,
            escaped_len(&self.value)?,
            escaped_len(&value)?,
        )?;
        ensure_tag_budget(wire_len)?;
        let previous = std::mem::replace(&mut self.value, value);
        self.wire_len = wire_len;
        Ok(previous)
    }

    /// Add one local namespace declaration.
    pub fn with_namespace(mut self, attr: raw::Attr) -> Result<Self> {
        check_namespace(&self.namespaces, &self.attrs, &attr, &["p", "xml"])?;
        let wire_len = checked_wire_add(self.wire_len, attr_wire_len(&attr)?)?;
        ensure_tag_budget(wire_len)?;
        self.namespaces.push(attr);
        self.wire_len = wire_len;
        Ok(self)
    }

    /// Add one prefixed inert extension attribute.
    pub fn with_attr(mut self, attr: raw::Attr) -> Result<Self> {
        check_extension(&self.attrs, &self.namespaces, &attr, &["p", "xml"])?;
        let wire_len = checked_wire_add(self.wire_len, attr_wire_len(&attr)?)?;
        ensure_tag_budget(wire_len)?;
        self.attrs.push(attr);
        self.wire_len = wire_len;
        Ok(self)
    }
}

/// A detached, source-ordered programmable-tag list.
///
/// The checked-in `[MS-OE376]` section 2.1.1170(c) states that PowerPoint
/// requires names within one `tagLst` to be case-insensitively unique. One
/// NFD/default-case-fold/NFD identity therefore drives lookup and every CRUD
/// operation while producer spelling remains unchanged. Parsing retains
/// malformed duplicates so callers can inspect and repair them by numeric
/// position; semantic lookup reports [`Error::AmbiguousName`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct List {
    tags: Vec<Tag>,
    namespaces: Vec<raw::Attr>,
    attrs: Vec<raw::Attr>,
    wire_len: usize,
}

impl List {
    /// Construct an empty list.
    pub fn new() -> Self {
        Self {
            tags: Vec::new(),
            namespaces: Vec::new(),
            attrs: Vec::new(),
            wire_len: EMPTY_WIRE_BYTES,
        }
    }

    /// Return all tags in source order.
    pub fn tags(&self) -> &[Tag] {
        &self.tags
    }

    /// Iterate without exposing a mutable slice that could violate uniqueness.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Tag> {
        self.tags.iter()
    }

    /// Return the number of tags.
    pub fn len(&self) -> usize {
        self.tags.len()
    }

    /// Report whether the list is empty.
    pub fn is_empty(&self) -> bool {
        self.tags.is_empty()
    }

    /// Return retained root namespace declarations.
    pub fn namespaces(&self) -> &[raw::Attr] {
        &self.namespaces
    }

    /// Return retained root extension attributes.
    pub fn attrs(&self) -> &[raw::Attr] {
        &self.attrs
    }

    /// Select by semantic name or checked numeric position.
    pub fn get<'a, 'k>(&'a self, key: impl Into<Key<'k>>) -> Result<&'a Tag> {
        let offset = self.offset(key.into())?;
        self.tags.get(offset).ok_or(Error::IndexOutOfBounds {
            index: offset,
            len: self.tags.len(),
        })
    }

    /// Append an owned tag without copying its strings.
    pub fn add(&mut self, tag: Tag) -> Result<()> {
        self.ensure_can_add(&tag, None)?;
        let wire_len = self.wire_after_add(&tag)?;
        self.tags.push(tag);
        self.wire_len = wire_len;
        Ok(())
    }

    /// Insert an owned tag at a checked source-order position.
    pub fn insert(&mut self, index: usize, tag: Tag) -> Result<()> {
        if index > self.tags.len() {
            return Err(Error::IndexOutOfBounds {
                index,
                len: self.tags.len(),
            });
        }
        self.ensure_can_add(&tag, None)?;
        let wire_len = self.wire_after_add(&tag)?;
        self.tags.insert(index, tag);
        self.wire_len = wire_len;
        Ok(())
    }

    /// Replace a selected tag by moving in its successor and returning the old
    /// allocation.
    pub fn replace<'k>(&mut self, key: impl Into<Key<'k>>, tag: Tag) -> Result<Tag> {
        let offset = self.offset(key.into())?;
        self.ensure_can_add(&tag, Some(offset))?;
        let previous_wire_len = self
            .tags
            .get(offset)
            .map(|previous| previous.wire_len)
            .ok_or(Error::IndexOutOfBounds {
                index: offset,
                len: self.tags.len(),
            })?;
        let wire_len = replace_wire_component(self.wire_len, previous_wire_len, tag.wire_len)?;
        ensure_list_budget(wire_len)?;
        let len = self.tags.len();
        let slot = self
            .tags
            .get_mut(offset)
            .ok_or(Error::IndexOutOfBounds { index: offset, len })?;
        let previous = std::mem::replace(slot, tag);
        self.wire_len = wire_len;
        Ok(previous)
    }

    /// Replace only a selected tag's inert value.
    pub fn set<'k>(&mut self, key: impl Into<Key<'k>>, value: impl Into<String>) -> Result<String> {
        let offset = self.offset(key.into())?;
        let value = value.into();
        bounded_text(&value, "tag value")?;
        let tag = self.tags.get(offset).ok_or(Error::IndexOutOfBounds {
            index: offset,
            len: self.tags.len(),
        })?;
        let tag_wire_len =
            replace_wire_component(tag.wire_len, escaped_len(&tag.value)?, escaped_len(&value)?)?;
        ensure_tag_budget(tag_wire_len)?;
        let wire_len = replace_wire_component(self.wire_len, tag.wire_len, tag_wire_len)?;
        ensure_list_budget(wire_len)?;
        let len = self.tags.len();
        let tag = self
            .tags
            .get_mut(offset)
            .ok_or(Error::IndexOutOfBounds { index: offset, len })?;
        let previous = std::mem::replace(&mut tag.value, value);
        tag.wire_len = tag_wire_len;
        self.wire_len = wire_len;
        Ok(previous)
    }

    /// Remove and return a selected tag without panicking on a stale selector.
    pub fn remove<'k>(&mut self, key: impl Into<Key<'k>>) -> Result<Tag> {
        let offset = self.offset(key.into())?;
        let removed_wire_len =
            self.tags
                .get(offset)
                .map(|tag| tag.wire_len)
                .ok_or(Error::IndexOutOfBounds {
                    index: offset,
                    len: self.tags.len(),
                })?;
        let wire_len = if self.tags.len() == 1 {
            let without_tag = checked_wire_sub(self.wire_len, removed_wire_len)?;
            replace_wire_component(without_tag, NONEMPTY_WIRE_BYTES, EMPTY_WIRE_BYTES)?
        } else {
            checked_wire_sub(self.wire_len, removed_wire_len)?
        };
        let mut current = 0usize;
        let removed = {
            let mut matches = self.tags.extract_if(.., |_| {
                let selected = current == offset;
                current = current.saturating_add(1);
                selected
            });
            matches.next()
        };
        let removed = removed.ok_or(Error::IndexOutOfBounds {
            index: offset,
            len: self.tags.len(),
        })?;
        self.wire_len = wire_len;
        Ok(removed)
    }

    /// Apply a complete checked permutation by names or numeric positions.
    ///
    /// `list.reorder(&["second", "first"])` is the semantic common path;
    /// `list.reorder(&[1_usize, 0])` supports source-order repair tooling.
    pub fn reorder<'k, K>(&mut self, order: &'k [K]) -> Result<()>
    where
        K: Copy + Into<Key<'k>>,
    {
        if order.len() != self.tags.len() {
            return Err(Error::OrderLength {
                expected: self.tags.len(),
                actual: order.len(),
            });
        }
        let mut ranks = vec![usize::MAX; self.tags.len()];
        for (rank, key) in order.iter().copied().enumerate() {
            let offset = self.offset(key.into())?;
            let len = ranks.len();
            let slot = ranks
                .get_mut(offset)
                .ok_or(Error::IndexOutOfBounds { index: offset, len })?;
            if *slot != usize::MAX {
                return Err(Error::DuplicateSelection { index: offset });
            }
            *slot = rank;
        }
        if let Some(index) = ranks.iter().position(|rank| *rank == usize::MAX) {
            return Err(invalid(format!(
                "physical tag at index {index} is absent from the reorder"
            )));
        }

        let tags = std::mem::take(&mut self.tags);
        let mut ranked = tags
            .into_iter()
            .zip(ranks)
            .map(|(tag, rank)| (rank, tag))
            .collect::<Vec<_>>();
        ranked.sort_by_key(|(rank, _)| *rank);
        self.tags = ranked.into_iter().map(|(_, tag)| tag).collect();
        Ok(())
    }

    /// Move all owned tags out of this list.
    pub fn into_tags(self) -> Vec<Tag> {
        self.tags
    }

    /// Add one root namespace declaration.
    pub fn with_namespace(mut self, attr: raw::Attr) -> Result<Self> {
        check_namespace(&self.namespaces, &self.attrs, &attr, &["p", "xml"])?;
        let wire_len = checked_wire_add(self.wire_len, attr_wire_len(&attr)?)?;
        ensure_list_budget(wire_len)?;
        self.namespaces.push(attr);
        self.wire_len = wire_len;
        Ok(self)
    }

    /// Add one prefixed inert root extension attribute.
    pub fn with_attr(mut self, attr: raw::Attr) -> Result<Self> {
        check_extension(&self.attrs, &self.namespaces, &attr, &["p", "xml"])?;
        let wire_len = checked_wire_add(self.wire_len, attr_wire_len(&attr)?)?;
        ensure_list_budget(wire_len)?;
        self.attrs.push(attr);
        self.wire_len = wire_len;
        Ok(self)
    }

    /// Encode with the requested namespace conformance.
    pub fn xml(&self, conformance: Conformance) -> Result<Vec<u8>> {
        write(self, conformance)
    }

    fn offset(&self, key: Key<'_>) -> Result<usize> {
        match key {
            Key::Name(name) => unique_offset(&self.tags, name),
            Key::Index(index) => {
                if index < self.tags.len() {
                    Ok(index)
                } else {
                    Err(Error::IndexOutOfBounds {
                        index,
                        len: self.tags.len(),
                    })
                }
            },
        }
    }

    fn ensure_can_add(&self, tag: &Tag, replaced: Option<usize>) -> Result<()> {
        validate_tag(tag)?;
        validate_tag_context(tag, &bound_prefixes(&self.namespaces, &["p", "xml"])?)?;
        if replaced.is_none() && self.tags.len() == MAX_TAGS {
            return Err(Error::Limit {
                resource: "tag count",
                limit: MAX_TAGS,
            });
        }
        let matches = self
            .tags
            .iter()
            .enumerate()
            .filter(|(index, existing)| {
                Some(*index) != replaced && same_name(existing.name(), tag.name())
            })
            .count();
        if matches == 0 {
            Ok(())
        } else {
            Err(Error::DuplicateName {
                name: tag.name.clone(),
                matches,
            })
        }
    }

    fn wire_after_add(&self, tag: &Tag) -> Result<usize> {
        let wire_len = if self.tags.is_empty() {
            let with_children =
                replace_wire_component(self.wire_len, EMPTY_WIRE_BYTES, NONEMPTY_WIRE_BYTES)?;
            checked_wire_add(with_children, tag.wire_len)?
        } else {
            checked_wire_add(self.wire_len, tag.wire_len)?
        };
        ensure_list_budget(wire_len)?;
        Ok(wire_len)
    }
}

impl Default for List {
    fn default() -> Self {
        Self::new()
    }
}

impl IntoIterator for List {
    type Item = Tag;
    type IntoIter = std::vec::IntoIter<Tag>;

    fn into_iter(self) -> Self::IntoIter {
        self.tags.into_iter()
    }
}

impl<'a> IntoIterator for &'a List {
    type Item = &'a Tag;
    type IntoIter = std::slice::Iter<'a, Tag>;

    fn into_iter(self) -> Self::IntoIter {
        self.tags.iter()
    }
}

/// One slide relationship source and its parsed detached list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Source {
    relationship_id: String,
    part_name: PackURI,
    list: List,
}

impl Source {
    /// Return the relationship ID on the source slide.
    pub fn rel(&self) -> &str {
        &self.relationship_id
    }

    /// Return the typed absolute target part name.
    pub fn part(&self) -> &PackURI {
        &self.part_name
    }

    /// Borrow the parsed list.
    pub fn list(&self) -> &List {
        &self.list
    }

    /// Move the parsed list out of its source descriptor.
    pub fn into_list(self) -> List {
        self.list
    }
}

/// Report whether a relationship type identifies a tag-list part.
pub fn is_relationship(value: &str) -> bool {
    matches!(value, TAG_REL | STRICT_TAG_REL)
}

/// Parse one bounded Strict or Transitional tag-list part.
pub fn parse(xml: &[u8]) -> Result<List> {
    if xml.len() > MAX_PART_BYTES {
        return Err(Error::Limit {
            resource: "tag-list bytes",
            limit: MAX_PART_BYTES,
        });
    }
    let xml = process_ooxml(xml)?;
    if xml.len() > MAX_PART_BYTES {
        return Err(Error::Limit {
            resource: "MCE-expanded tag-list bytes",
            limit: MAX_PART_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root = false;
    let mut closed = false;
    let mut open_tag: Option<(usize, Tag)> = None;
    let mut tags = Vec::new();
    let mut namespaces = Vec::new();
    let mut attrs = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let name = element.local_name();
                if !root && depth == 0 && pml(&namespace) && name.as_ref() == b"tagLst" {
                    root = true;
                    let parsed = parse_attributes(&element, &[], reader.decoder())?;
                    namespaces = parsed.namespaces;
                    attrs = parsed.extensions;
                    depth = 1;
                } else if root
                    && !closed
                    && depth == 1
                    && open_tag.is_none()
                    && pml(&namespace)
                    && name.as_ref() == b"tag"
                {
                    if tags.len() == MAX_TAGS {
                        return Err(Error::Limit {
                            resource: "tag count",
                            limit: MAX_TAGS,
                        });
                    }
                    let tag = parse_tag(&element, reader.decoder())?;
                    depth = depth.saturating_add(1);
                    open_tag = Some((depth, tag));
                } else {
                    return Err(invalid(format!(
                        "unexpected tag-list element '{}'",
                        String::from_utf8_lossy(name.as_ref())
                    )));
                }
            },
            Event::Empty(element) => {
                let name = element.local_name();
                if !root && depth == 0 && pml(&namespace) && name.as_ref() == b"tagLst" {
                    let parsed = parse_attributes(&element, &[], reader.decoder())?;
                    namespaces = parsed.namespaces;
                    attrs = parsed.extensions;
                    root = true;
                    closed = true;
                } else if root
                    && !closed
                    && depth == 1
                    && open_tag.is_none()
                    && pml(&namespace)
                    && name.as_ref() == b"tag"
                {
                    if tags.len() == MAX_TAGS {
                        return Err(Error::Limit {
                            resource: "tag count",
                            limit: MAX_TAGS,
                        });
                    }
                    tags.push(parse_tag(&element, reader.decoder())?);
                } else {
                    return Err(invalid(format!(
                        "unexpected tag-list element '{}'",
                        String::from_utf8_lossy(name.as_ref())
                    )));
                }
            },
            Event::End(element) => {
                let name = element.local_name();
                if open_tag.as_ref().is_some_and(|(level, _)| *level == depth)
                    && pml(&namespace)
                    && name.as_ref() == b"tag"
                {
                    let Some((_, tag)) = open_tag.take() else {
                        return Err(invalid("tag parser state is inconsistent"));
                    };
                    tags.push(tag);
                    depth = depth.saturating_sub(1);
                } else if root
                    && !closed
                    && depth == 1
                    && pml(&namespace)
                    && name.as_ref() == b"tagLst"
                {
                    closed = true;
                    depth = 0;
                } else {
                    return Err(invalid("unexpected tag-list end element"));
                }
            },
            Event::Text(text) => {
                let value = text.decode().map_err(xml_error)?;
                let value = quick_xml::escape::unescape(&value).map_err(xml_error)?;
                if !value.trim().is_empty() {
                    return Err(invalid("tag elements cannot contain text"));
                }
            },
            Event::CData(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("tag elements cannot contain CDATA"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTDs and processing instructions are rejected in tag lists",
                ));
            },
            Event::GeneralRef(_) => {
                return Err(invalid("tag elements cannot contain entity references"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !root || !closed || depth != 0 || open_tag.is_some() {
        return Err(invalid("unterminated tag-list part"));
    }
    let wire_len = list_wire_len_parts(&tags, &namespaces, &attrs)?;
    let list = List {
        tags,
        namespaces,
        attrs,
        wire_len,
    };
    validate_structure(&list)?;
    Ok(list)
}

/// Encode one detached list without interpreting any retained value.
pub fn write(value: &List, conformance: Conformance) -> Result<Vec<u8>> {
    validate_structure(value)?;
    validate_unique_names(value)?;
    ensure_list_budget(value.wire_len)?;
    let mut out = Vec::with_capacity(value.wire_len);
    append(&mut out, XML_DECL)?;
    append(&mut out, ROOT_OPEN)?;
    escape(&mut out, conformance.namespace())?;
    push(&mut out, b'\"')?;
    for attr in &value.namespaces {
        write_preserved(&mut out, attr)?;
    }
    for attr in &value.attrs {
        write_preserved(&mut out, attr)?;
    }
    if value.tags.is_empty() {
        append(&mut out, ROOT_EMPTY_CLOSE)?;
        return Ok(out);
    }
    append(&mut out, ROOT_CHILDREN_OPEN)?;
    for tag in &value.tags {
        append(&mut out, TAG_OPEN)?;
        for attr in &tag.namespaces {
            write_preserved(&mut out, attr)?;
        }
        for attr in &tag.attrs {
            write_preserved(&mut out, attr)?;
        }
        write_attr(&mut out, "name", &tag.name)?;
        write_attr(&mut out, "val", &tag.value)?;
        append(&mut out, TAG_CLOSE)?;
    }
    append(&mut out, ROOT_CLOSE)?;
    Ok(out)
}

/// Discover and parse every internal tag-list relationship on one slide part.
///
/// OPC relationship storage does not retain XML source order. Results are
/// therefore returned in ascending relationship-ID byte order, matching the
/// package writer's stable relationship serialization order.
pub fn discover(slide: &dyn OpcPart, package: &OpcPackage) -> Result<Vec<Source>> {
    let mut relationships = slide
        .rels()
        .iter()
        .filter(|relationship| is_relationship(relationship.reltype()))
        .collect::<Vec<_>>();
    if relationships.len() > MAX_TAG_PARTS {
        return Err(Error::Limit {
            resource: "slide tag-list relationships",
            limit: MAX_TAG_PARTS,
        });
    }
    relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    let mut targets = HashSet::with_capacity(relationships.len());
    let mut output = Vec::with_capacity(relationships.len());
    for relationship in relationships {
        if relationship.is_external() {
            return Err(invalid(format!(
                "tag-list relationship '{}' cannot be external",
                relationship.r_id()
            )));
        }
        let target = relationship.target_partname()?;
        if !targets.insert(target.to_string()) {
            return Err(invalid(format!(
                "duplicate slide tag-list target '{target}'"
            )));
        }
        let part = package.get_part(&target)?;
        if part.content_type() != CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: CONTENT_TYPE.into(),
                actual: part.content_type().into(),
            });
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid(format!(
                "tag-list part '{target}' has unexpected relationships"
            )));
        }
        output.push(Source {
            relationship_id: relationship.r_id().into(),
            part_name: target,
            list: parse(part.blob())?,
        });
    }
    Ok(output)
}

struct ParsedAttributes {
    values: Vec<(String, String)>,
    namespaces: Vec<raw::Attr>,
    extensions: Vec<raw::Attr>,
}

fn parse_attributes(
    element: &BytesStart<'_>,
    known: &[&str],
    decoder: Decoder,
) -> Result<ParsedAttributes> {
    let mut values = Vec::new();
    let mut namespaces = Vec::new();
    let mut extensions = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(xml_error)?
            .to_string();
        validate_qname(&name)?;
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded_text(&value, "tag attribute")?;
        if name == "xmlns:p" {
            if !matches!(value.as_str(), PML_TEXT | STRICT_TEXT) {
                return Err(invalid(
                    "p prefix is bound to a non-PresentationML namespace",
                ));
            }
        } else if is_namespace(&name) {
            namespaces.push(raw::Attr {
                qualified_name: name,
                value,
            });
        } else if !name.contains(':') && known.contains(&name.as_str()) {
            values.push((name, value));
        } else if name.contains(':') {
            extensions.push(raw::Attr {
                qualified_name: name,
                value,
            });
        } else {
            return Err(invalid(format!("unexpected tag attribute '{name}'")));
        }
    }
    Ok(ParsedAttributes {
        values,
        namespaces,
        extensions,
    })
}

fn parse_tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let parsed = parse_attributes(element, &["name", "val"], decoder)?;
    let mut name = None;
    let mut value = None;
    for (key, item) in parsed.values {
        match key.as_str() {
            "name" => name = Some(item),
            "val" => value = Some(item),
            _ => return Err(invalid("unexpected parsed tag attribute")),
        }
    }
    let name = name.ok_or_else(|| invalid("tag is missing 'name'"))?;
    let value = value.ok_or_else(|| invalid("tag is missing 'val'"))?;
    let wire_len = tag_wire_len_parts(&name, &value, &parsed.namespaces, &parsed.extensions)?;
    Ok(Tag {
        name,
        value,
        namespaces: parsed.namespaces,
        attrs: parsed.extensions,
        wire_len,
    })
}

fn unique_offset(tags: &[Tag], name: &str) -> Result<usize> {
    let mut first = None;
    let mut matches = 0usize;
    for (index, tag) in tags.iter().enumerate() {
        if same_name(tag.name(), name) {
            first = first.or(Some(index));
            matches = matches.saturating_add(1);
        }
    }
    match (first, matches) {
        (Some(index), 1) => Ok(index),
        (Some(_), count) => Err(Error::AmbiguousName {
            name: name.into(),
            matches: count,
        }),
        (None, _) => Err(Error::NameNotFound(name.into())),
    }
}

fn same_name(left: &str, right: &str) -> bool {
    name_key(left) == name_key(right)
}

fn name_key(value: &str) -> String {
    value.chars().nfd().default_case_fold().nfd().collect()
}

fn escaped_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let bytes = match character {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        total
            .checked_add(bytes)
            .ok_or_else(|| invalid("escaped tag text length overflow"))
    })
}

fn attr_wire_len(attr: &raw::Attr) -> Result<usize> {
    attr_wire_len_parts(attr.qualified_name(), attr.value())
}

fn attr_wire_len_parts(name: &str, value: &str) -> Result<usize> {
    checked_wire_add(
        name.len()
            .checked_add(4)
            .ok_or_else(|| invalid("tag attribute length overflow"))?,
        escaped_len(value)?,
    )
}

fn tag_wire_len_parts(
    name: &str,
    value: &str,
    namespaces: &[raw::Attr],
    attrs: &[raw::Attr],
) -> Result<usize> {
    let mut wire_len = TAG_OPEN.len() + TAG_CLOSE.len();
    for attr in namespaces.iter().chain(attrs) {
        wire_len = checked_wire_add(wire_len, attr_wire_len(attr)?)?;
    }
    wire_len = checked_wire_add(wire_len, attr_wire_len_parts("name", name)?)?;
    checked_wire_add(wire_len, attr_wire_len_parts("val", value)?)
}

fn list_wire_len_parts(
    tags: &[Tag],
    namespaces: &[raw::Attr],
    attrs: &[raw::Attr],
) -> Result<usize> {
    let mut wire_len = if tags.is_empty() {
        EMPTY_WIRE_BYTES
    } else {
        NONEMPTY_WIRE_BYTES
    };
    for attr in namespaces.iter().chain(attrs) {
        wire_len = checked_wire_add(wire_len, attr_wire_len(attr)?)?;
    }
    for tag in tags {
        wire_len = checked_wire_add(wire_len, tag.wire_len)?;
    }
    Ok(wire_len)
}

fn checked_wire_add(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid("tag-list wire length overflow"))
}

fn checked_wire_sub(left: usize, right: usize) -> Result<usize> {
    left.checked_sub(right)
        .ok_or_else(|| invalid("tag-list wire length invariant underflow"))
}

fn replace_wire_component(current: usize, old: usize, new: usize) -> Result<usize> {
    checked_wire_add(checked_wire_sub(current, old)?, new)
}

fn ensure_tag_budget(wire_len: usize) -> Result<()> {
    ensure_list_budget(checked_wire_add(NONEMPTY_WIRE_BYTES, wire_len)?)
}

fn ensure_list_budget(wire_len: usize) -> Result<()> {
    if wire_len <= MAX_PART_BYTES {
        Ok(())
    } else {
        Err(Error::Limit {
            resource: "encoded tag-list bytes",
            limit: MAX_PART_BYTES,
        })
    }
}

fn validate_structure(value: &List) -> Result<()> {
    if value.tags.len() > MAX_TAGS {
        return Err(Error::Limit {
            resource: "tag count",
            limit: MAX_TAGS,
        });
    }
    validate_element_attrs(&value.namespaces, &value.attrs, &["xmlns:p"])?;
    let root_prefixes = bound_prefixes(&value.namespaces, &["p", "xml"])?;
    validate_bound_attrs(&value.attrs, &root_prefixes)?;
    for tag in &value.tags {
        validate_tag(tag)?;
        validate_tag_context(tag, &root_prefixes)?;
    }
    let wire_len = list_wire_len_parts(&value.tags, &value.namespaces, &value.attrs)?;
    if wire_len != value.wire_len {
        return Err(invalid("tag-list wire-size invariant is inconsistent"));
    }
    ensure_list_budget(wire_len)?;
    Ok(())
}

fn validate_unique_names(value: &List) -> Result<()> {
    let mut seen = HashMap::<String, usize>::with_capacity(value.tags.len());
    for tag in &value.tags {
        let count = seen.entry(name_key(tag.name())).or_default();
        if *count != 0 {
            return Err(Error::DuplicateName {
                name: tag.name.clone(),
                matches: *count,
            });
        }
        *count = count.saturating_add(1);
    }
    Ok(())
}

fn validate_tag(tag: &Tag) -> Result<()> {
    bounded_text(&tag.name, "tag name")?;
    bounded_text(&tag.value, "tag value")?;
    for attr in tag.namespaces.iter().chain(&tag.attrs) {
        validate_qname(&attr.qualified_name)?;
        bounded_text(&attr.value, "tag attribute")?;
    }
    let wire_len = tag_wire_len_parts(&tag.name, &tag.value, &tag.namespaces, &tag.attrs)?;
    if wire_len != tag.wire_len {
        return Err(invalid("tag wire-size invariant is inconsistent"));
    }
    Ok(())
}

fn validate_tag_context(tag: &Tag, root_prefixes: &[String]) -> Result<()> {
    validate_element_attrs(&tag.namespaces, &tag.attrs, &["name", "val"])?;
    let prefixes = bound_prefixes(&tag.namespaces, root_prefixes)?;
    validate_bound_attrs(&tag.attrs, &prefixes)
}

fn validate_element_attrs(
    namespaces: &[raw::Attr],
    attrs: &[raw::Attr],
    reserved: &[&str],
) -> Result<()> {
    let mut seen = reserved.iter().copied().collect::<HashSet<_>>();
    for attr in namespaces.iter().chain(attrs) {
        if !seen.insert(attr.qualified_name()) {
            return Err(invalid(format!(
                "duplicate tag attribute '{}'",
                attr.qualified_name()
            )));
        }
    }
    Ok(())
}

fn bound_prefixes<S: AsRef<str>>(namespaces: &[raw::Attr], inherited: &[S]) -> Result<Vec<String>> {
    let mut prefixes = inherited
        .iter()
        .map(|prefix| prefix.as_ref().to_owned())
        .collect::<Vec<_>>();
    for attr in namespaces {
        let Some(prefix) = attr.qualified_name().strip_prefix("xmlns:") else {
            continue;
        };
        if prefix == "xmlns" || prefix == "p" {
            return Err(invalid(format!(
                "namespace prefix '{prefix}' is reserved in a tag list"
            )));
        }
        if prefix == "xml" && attr.value() != "http://www.w3.org/XML/1998/namespace" {
            return Err(invalid("xml prefix has an invalid namespace binding"));
        }
        if attr.value().is_empty() {
            return Err(invalid("prefixed namespace URI cannot be empty"));
        }
        if !prefixes.iter().any(|existing| existing == prefix) {
            prefixes.push(prefix.into());
        }
    }
    Ok(prefixes)
}

fn validate_bound_attrs(attrs: &[raw::Attr], prefixes: &[String]) -> Result<()> {
    for attr in attrs {
        let Some((prefix, _)) = attr.qualified_name().split_once(':') else {
            return Err(invalid("extension tag attributes must be prefixed"));
        };
        if !prefixes.iter().any(|bound| bound == prefix) {
            return Err(invalid(format!(
                "tag attribute prefix '{prefix}' has no namespace binding"
            )));
        }
    }
    Ok(())
}

fn check_namespace(
    target: &[raw::Attr],
    attrs: &[raw::Attr],
    attr: &raw::Attr,
    inherited: &[&str],
) -> Result<()> {
    if !is_namespace(attr.qualified_name()) {
        return Err(invalid("tag namespace must be named xmlns or xmlns:prefix"));
    }
    if attr.qualified_name() == "xmlns:p" {
        return Err(invalid("p namespace is selected by Conformance"));
    }
    if target
        .iter()
        .any(|existing| existing.qualified_name() == attr.qualified_name())
    {
        return Err(invalid(format!(
            "duplicate namespace declaration '{}'",
            attr.qualified_name()
        )));
    }
    let prefixes = bound_prefixes(target, inherited)?;
    let prefixes = bound_prefixes(std::slice::from_ref(attr), &prefixes)?;
    validate_bound_attrs(attrs, &prefixes)?;
    Ok(())
}

fn check_extension(
    target: &[raw::Attr],
    namespaces: &[raw::Attr],
    attr: &raw::Attr,
    inherited: &[&str],
) -> Result<()> {
    if is_namespace(attr.qualified_name()) || !attr.qualified_name().contains(':') {
        return Err(invalid("extension tag attributes must be prefixed"));
    }
    if target
        .iter()
        .any(|existing| existing.qualified_name() == attr.qualified_name())
    {
        return Err(invalid(format!(
            "duplicate extension attribute '{}'",
            attr.qualified_name()
        )));
    }
    let prefixes = bound_prefixes(namespaces, inherited)?;
    validate_bound_attrs(std::slice::from_ref(attr), &prefixes)?;
    Ok(())
}

fn write_preserved(out: &mut Vec<u8>, attr: &raw::Attr) -> Result<()> {
    validate_qname(attr.qualified_name())?;
    bounded_text(attr.value(), "tag attribute")?;
    write_attr(out, attr.qualified_name(), attr.value())
}

fn write_attr(out: &mut Vec<u8>, name: &str, value: &str) -> Result<()> {
    push(out, b' ')?;
    append(out, name.as_bytes())?;
    append(out, b"=\"")?;
    escape(out, value)?;
    push(out, b'\"')
}

fn escape(out: &mut Vec<u8>, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => append(out, b"&amp;")?,
            '<' => append(out, b"&lt;")?,
            '"' => append(out, b"&quot;")?,
            '\t' => append(out, b"&#x9;")?,
            '\n' => append(out, b"&#xA;")?,
            '\r' => append(out, b"&#xD;")?,
            _ => {
                let mut encoded = [0; 4];
                append(out, character.encode_utf8(&mut encoded).as_bytes())?;
            },
        }
    }
    Ok(())
}

fn append(out: &mut Vec<u8>, bytes: &[u8]) -> Result<()> {
    let len = out
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("tag-list output length overflow"))?;
    if len > MAX_PART_BYTES {
        return Err(Error::Limit {
            resource: "encoded tag-list bytes",
            limit: MAX_PART_BYTES,
        });
    }
    out.extend_from_slice(bytes);
    Ok(())
}

fn push(out: &mut Vec<u8>, byte: u8) -> Result<()> {
    append(out, std::slice::from_ref(&byte))
}

fn is_namespace(value: &str) -> bool {
    value == "xmlns" || value.starts_with("xmlns:")
}

fn validate_qname(value: &str) -> Result<()> {
    let mut parts = value.split(':');
    let Some(first) = parts.next() else {
        return Err(invalid("tag attribute QName is empty"));
    };
    let second = parts.next();
    let valid = valid_ncname(first)
        && second.is_none_or(valid_ncname)
        && parts.next().is_none()
        && value.len() <= MAX_TEXT_BYTES;
    if valid {
        Ok(())
    } else {
        Err(invalid(format!("invalid tag attribute QName '{value}'")))
    }
}

fn valid_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first == '_' || first.is_alphabetic())
        && chars.all(|character| {
            character == '_'
                || character == '-'
                || character == '.'
                || character == '\u{B7}'
                || character.is_alphanumeric()
                || matches!(character, '\u{0300}'..='\u{036F}' | '\u{203F}'..='\u{2040}')
        })
}

fn bounded_text(value: &str, resource: &'static str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(Error::Limit {
            resource,
            limit: MAX_TEXT_BYTES,
        });
    }
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid(format!(
            "{resource} contains a character forbidden by XML 1.0"
        )))
    }
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

fn pml(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == PML || value.as_ref() == STRICT)
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    #[test]
    fn strict_round_trip_preserves_inert_values_and_extensions() {
        let xml = br#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:x" x:origin="fixture"><p:tag name="PATH" val="C:\Docs\file"/><p:tag name="XML" val="&lt;root command=&quot;none&quot;/&gt;"></p:tag></p:tagLst>"#;
        let value = parse(xml).unwrap();
        assert_eq!(
            value.get("xml").unwrap().value(),
            "<root command=\"none\"/>"
        );
        assert_eq!(value.namespaces()[0].qualified_name(), "xmlns:x");
        assert_eq!(value.attrs()[0].qualified_name(), "x:origin");

        let strict = value.xml(Conformance::Strict).unwrap();
        assert!(std::str::from_utf8(&strict).unwrap().contains(STRICT_TEXT));
        assert_eq!(parse(&strict).unwrap(), value);
    }

    #[test]
    fn mce_fallback_is_selected() {
        let xml = br#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:tag/></mc:Choice><mc:Fallback><p:tag name="fallback" val="1"/></mc:Fallback></mc:AlternateContent></p:tagLst>"#;
        assert_eq!(parse(xml).unwrap().get("FALLBACK").unwrap().value(), "1");
    }

    #[test]
    fn unicode_caseless_crud_is_checked_and_move_first() {
        let mut list = List::new();
        list.add(Tag::new("Straße", "one").unwrap()).unwrap();
        list.add(Tag::new("Owner", "Alice").unwrap()).unwrap();
        assert_eq!(list.get("STRASSE").unwrap().value(), "one");
        assert_eq!(list.get(1_usize).unwrap().name(), "Owner");
        assert!(matches!(
            list.get(2_usize),
            Err(Error::IndexOutOfBounds { index: 2, len: 2 })
        ));
        assert!(matches!(
            list.add(Tag::new("strasse", "duplicate").unwrap()),
            Err(Error::DuplicateName { matches: 1, .. })
        ));

        let old = list
            .replace("OWNER", Tag::new("Reviewer", "Bob").unwrap())
            .unwrap();
        assert_eq!(old.value(), "Alice");
        assert_eq!(list.set("reviewer", "Carol").unwrap(), "Bob");
        list.insert(1, Tag::new("Status", "Draft").unwrap())
            .unwrap();
        list.reorder(&["status", "strasse", "reviewer"]).unwrap();
        assert_eq!(list.tags()[0].name(), "Status");
        list.reorder(&[2_usize, 1, 0]).unwrap();
        assert_eq!(list.tags()[0].name(), "Reviewer");
        assert_eq!(list.remove("STATUS").unwrap().value(), "Draft");
    }

    #[test]
    fn malformed_duplicate_names_have_typed_ambiguity_and_numeric_repair() {
        let xml = format!(
            r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="Straße" val="A"/><p:tag name="STRASSE" val="B"/></p:tagLst>"#
        );
        let mut list = parse(xml.as_bytes()).unwrap();
        assert!(matches!(
            list.get("strasse"),
            Err(Error::AmbiguousName { matches: 2, .. })
        ));
        assert!(matches!(
            write(&list, Conformance::Transitional),
            Err(Error::DuplicateName { matches: 1, .. })
        ));
        assert_eq!(list.remove(1_usize).unwrap().value(), "B");
        assert_eq!(list.get("STRASSE").unwrap().value(), "A");
        assert!(write(&list, Conformance::Transitional).is_ok());
    }

    #[test]
    fn reorder_rejects_partial_and_duplicate_orders_without_mutating() {
        let mut list = List::new();
        list.add(Tag::new("one", "1").unwrap()).unwrap();
        list.add(Tag::new("two", "2").unwrap()).unwrap();
        let original = list.clone();
        assert!(matches!(
            list.reorder(&["one"]),
            Err(Error::OrderLength {
                expected: 2,
                actual: 1
            })
        ));
        assert_eq!(list, original);
        assert!(matches!(
            list.reorder(&["one", "ONE"]),
            Err(Error::DuplicateSelection { index: 0 })
        ));
        assert_eq!(list, original);
    }

    #[test]
    fn malformed_markup_and_resource_limits_are_rejected() {
        for xml in [
            format!(r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag val="x"/></p:tagLst>"#),
            format!(r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="x"/></p:tagLst>"#),
            format!(
                r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="x" val="y"><p:tag name="z" val="q"/></p:tag></p:tagLst>"#
            ),
            format!(r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:other/></p:tagLst>"#),
            format!(r#"<!DOCTYPE x><p:tagLst xmlns:p="{TRANSITIONAL}"/>"#),
            format!(r#"<?bad x?><p:tagLst xmlns:p="{TRANSITIONAL}"/>"#),
        ] {
            assert!(parse(xml.as_bytes()).is_err(), "{xml}");
        }
        assert!(matches!(
            parse(&vec![b' '; MAX_PART_BYTES + 1]),
            Err(Error::Limit { .. })
        ));
        assert!(Tag::new("bad\0name", "value").is_err());

        let entity = format!(
            r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="x" val="y">&amp;</p:tag></p:tagLst>"#
        );
        assert!(matches!(
            parse(entity.as_bytes()),
            Err(Error::Invalid(message)) if message.contains("entity references")
        ));
    }

    #[test]
    fn parsing_rejects_canonical_escape_expansion_past_the_wire_budget() {
        const ENTITY: &str = "&#9;";
        let references_per_tag = (MAX_PART_BYTES - 512) / (2 * ENTITY.len());
        assert!(references_per_tag <= MAX_TEXT_BYTES);
        let references = ENTITY.repeat(references_per_tag);
        let xml = format!(
            r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="one" val="{references}"/><p:tag name="two" val="{references}"/></p:tagLst>"#
        );
        assert!(xml.len() <= MAX_PART_BYTES);
        assert!(matches!(
            parse(xml.as_bytes()),
            Err(Error::Limit {
                resource: "encoded tag-list bytes",
                ..
            })
        ));
    }

    #[test]
    fn discovery_uses_stable_relationship_id_order() {
        use litchi_opc::{Part, XmlPart};

        let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut slide = XmlPart::new(
            slide_name.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            Vec::new(),
        );
        for (relationship_id, target) in [("zId", "../tags/tag2.xml"), ("aId", "../tags/tag1.xml")]
        {
            slide.rels_mut().add_relationship(
                TAG_REL.into(),
                target.into(),
                relationship_id.into(),
                false,
            );
        }

        let mut package = OpcPackage::new();
        package.add_part(Box::new(slide));
        for (part_name, name) in [
            ("/ppt/tags/tag1.xml", "first"),
            ("/ppt/tags/tag2.xml", "second"),
        ] {
            package.add_part(Box::new(XmlPart::new(
                PackURI::new(part_name).unwrap(),
                CONTENT_TYPE.into(),
                format!(
                    r#"<p:tagLst xmlns:p="{TRANSITIONAL}"><p:tag name="{name}" val="1"/></p:tagLst>"#
                )
                .into_bytes(),
            )));
        }

        let source = package.get_part(&slide_name).unwrap();
        let discovered = discover(source, &package).unwrap();
        assert_eq!(
            discovered.iter().map(Source::rel).collect::<Vec<_>>(),
            ["aId", "zId"]
        );
        assert_eq!(discovered[0].list().get("first").unwrap().value(), "1");
        assert_eq!(discovered[1].list().get("second").unwrap().value(), "1");
    }

    #[test]
    fn raw_attributes_require_valid_bound_prefixes() {
        let unbound = raw::Attr::new("x:value", "1").unwrap();
        assert!(
            Tag::new("name", "value")
                .unwrap()
                .with_attr(unbound)
                .is_err()
        );

        let tag = Tag::new("name", "value")
            .unwrap()
            .with_namespace(raw::Attr::new("xmlns:x", "urn:x").unwrap())
            .unwrap()
            .with_attr(raw::Attr::new("x:value", "1").unwrap())
            .unwrap();
        let mut list = List::new();
        list.add(tag).unwrap();
        assert!(
            std::str::from_utf8(&write(&list, Conformance::Transitional).unwrap())
                .unwrap()
                .contains("x:value=\"1\"")
        );
    }

    #[test]
    fn escaped_size_budget_is_cached_and_failed_edits_are_atomic() {
        let quotes = "\"".repeat(MAX_TEXT_BYTES);
        assert_eq!(escaped_len(&quotes).unwrap(), 6 * MAX_TEXT_BYTES);

        let mut list = List::new()
            .with_namespace(raw::Attr::new("xmlns:x", "urn:x").unwrap())
            .unwrap()
            .with_attr(raw::Attr::new("x:padding", quotes.clone()).unwrap())
            .unwrap();
        list.add(Tag::new("small", "ok").unwrap()).unwrap();
        assert_eq!(
            list.wire_len,
            write(&list, Conformance::Transitional).unwrap().len()
        );

        {
            let before = list.clone();
            let replacement = Tag::new("large", quotes.clone()).unwrap();
            assert!(matches!(
                list.replace("small", replacement),
                Err(Error::Limit {
                    resource: "encoded tag-list bytes",
                    ..
                })
            ));
            assert_eq!(list, before);
        }
        {
            let before = list.clone();
            assert!(matches!(
                list.set("small", quotes.clone()),
                Err(Error::Limit {
                    resource: "encoded tag-list bytes",
                    ..
                })
            ));
            assert_eq!(list, before);
            assert_eq!(list.get("small").unwrap().value(), "ok");
        }
        {
            let before = list.clone();
            assert!(matches!(
                list.add(Tag::new("large", quotes.clone()).unwrap()),
                Err(Error::Limit {
                    resource: "encoded tag-list bytes",
                    ..
                })
            ));
            assert_eq!(list, before);
        }
        {
            let before = list.clone();
            assert!(matches!(
                list.insert(0, Tag::new("large", quotes.clone()).unwrap()),
                Err(Error::Limit {
                    resource: "encoded tag-list bytes",
                    ..
                })
            ));
            assert_eq!(list, before);
        }
        assert!(write(&list, Conformance::Transitional).is_ok());

        let root_overflow = List::new()
            .with_namespace(raw::Attr::new("xmlns:x", "urn:x").unwrap())
            .unwrap()
            .with_attr(raw::Attr::new("x:first", quotes.clone()).unwrap())
            .unwrap()
            .with_attr(raw::Attr::new("x:second", quotes.clone()).unwrap());
        assert!(matches!(root_overflow, Err(Error::Limit { .. })));

        let tag_overflow = Tag::new("standalone", "ok")
            .unwrap()
            .with_namespace(raw::Attr::new("xmlns:x", "urn:x").unwrap())
            .unwrap()
            .with_attr(raw::Attr::new("x:first", quotes.clone()).unwrap())
            .unwrap()
            .with_attr(raw::Attr::new("x:second", quotes).unwrap());
        assert!(matches!(tag_overflow, Err(Error::Limit { .. })));
    }

    #[test]
    fn namespace_builders_reject_invalid_prospective_bindings() {
        let tag = Tag::new("name", "value")
            .unwrap()
            .with_attr(raw::Attr::new("xml:lang", "en").unwrap())
            .unwrap();
        assert!(
            tag.with_namespace(
                raw::Attr::new("xmlns:xml", "https://example.invalid/not-xml").unwrap()
            )
            .is_err()
        );
        assert!(
            List::new()
                .with_namespace(raw::Attr::new("xmlns:xmlns", "urn:invalid").unwrap())
                .is_err()
        );
        assert!(
            Tag::new("name", "value")
                .unwrap()
                .with_namespace(raw::Attr::new("xmlns:x", "").unwrap())
                .is_err()
        );
    }
}
