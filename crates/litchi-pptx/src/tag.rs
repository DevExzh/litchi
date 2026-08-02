//! Inert PresentationML programmable tags and owner relationship discovery.
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
use litchi_opc::{OpcPackage, PackURI, Part as OpcPart, XmlPart};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::ops::Range;
use unicode_normalization::UnicodeNormalization;

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const PML_TEXT: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_TEXT: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const TAG_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
const STRICT_TAG_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/tags";
const REL_TEXT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_TEXT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

/// Content type of a PresentationML programmable-tag part.
pub const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";

const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_TAGS: usize = 16_384;
const MAX_TAG_PARTS: usize = 1_024;
const MAX_OWNER_BYTES: usize = 64 * 1024 * 1024;
const MAX_GRAPH_PARTS: usize = 100_000;
const MAX_GRAPH_LINKS: usize = 1_000_000;
const MAX_PART_NAME_BYTES: usize = 64 * 1024 * 1024;
const PART_NAME_ATTEMPTS: usize = 10_000;
const MAX_SOURCE_RELATIONSHIPS: usize = 65_536;
const MAX_OWNER_NODES: usize = 1_000_000;
const MAX_OWNER_DEPTH: usize = 512;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4_096;
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

    fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => TAG_REL,
            Self::Strict => STRICT_TAG_REL,
        }
    }

    fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => REL_TEXT,
            Self::Strict => STRICT_REL_TEXT,
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

/// One owner-part relationship source and its parsed detached list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Source {
    relationship_id: String,
    part_name: PackURI,
    conformance: Conformance,
    list: List,
}

impl Source {
    /// Return the relationship ID on the source PresentationML part.
    pub fn rel(&self) -> &str {
        &self.relationship_id
    }

    /// Return the typed absolute target part name.
    pub fn part(&self) -> &PackURI {
        &self.part_name
    }

    /// Return the namespace profile detected from the source tag-list root.
    pub fn conformance(&self) -> Conformance {
        self.conformance
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
    parse_profiled(xml).map(|(list, _)| list)
}

fn parse_profiled(xml: &[u8]) -> Result<(List, Conformance)> {
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
    let mut conformance = None;
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
                if !root && depth == 0 && pml(&namespace).is_some() && name.as_ref() == b"tagLst" {
                    root = true;
                    conformance = pml(&namespace);
                    let parsed = parse_attributes(&element, &[], reader.decoder())?;
                    namespaces = parsed.namespaces;
                    attrs = parsed.extensions;
                    depth = 1;
                } else if root
                    && !closed
                    && depth == 1
                    && open_tag.is_none()
                    && pml(&namespace) == conformance
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
                if !root && depth == 0 && pml(&namespace).is_some() && name.as_ref() == b"tagLst" {
                    conformance = pml(&namespace);
                    let parsed = parse_attributes(&element, &[], reader.decoder())?;
                    namespaces = parsed.namespaces;
                    attrs = parsed.extensions;
                    root = true;
                    closed = true;
                } else if root
                    && !closed
                    && depth == 1
                    && open_tag.is_none()
                    && pml(&namespace) == conformance
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
                    && pml(&namespace) == conformance
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
                    && pml(&namespace) == conformance
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
    let conformance =
        conformance.ok_or_else(|| invalid("tag-list namespace profile is missing"))?;
    Ok((list, conformance))
}

/// Encode one detached list without interpreting any retained value.
pub fn write(value: &List, conformance: Conformance) -> Result<Vec<u8>> {
    validate_structure(value)?;
    validate_unique_names(value)?;
    ensure_list_budget(value.wire_len)?;
    let mut out = Vec::new();
    out.try_reserve_exact(value.wire_len)
        .map_err(|source| allocation("encoded tag-list output", source))?;
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

/// Discover and parse every internal tag-list relationship on one owner part.
///
/// This is deliberately low-level diagnostic inventory: it does not inspect
/// XML anchors, so its results can include shape-owned and unanchored parts.
/// Use [`load`] for the part-level semantic attachment. OPC relationship
/// storage does not retain XML source order, so results are returned in
/// ascending relationship-ID byte order.
pub fn discover(owner: &dyn OpcPart, package: &OpcPackage) -> Result<Vec<Source>> {
    let mut scanned = 0usize;
    let mut relationships = Vec::new();
    relationships
        .try_reserve_exact(owner.rels().len().min(MAX_TAG_PARTS))
        .map_err(|source| allocation("tag relationship inventory", source))?;
    for relationship in owner.rels().iter() {
        bump_graph_link(&mut scanned)?;
        if !is_relationship(relationship.reltype()) {
            continue;
        }
        if relationships.len() == MAX_TAG_PARTS {
            return Err(Error::Limit {
                resource: "owner tag-list relationships",
                limit: MAX_TAG_PARTS,
            });
        }
        relationships.push(relationship);
    }
    relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    let mut targets = HashSet::new();
    targets
        .try_reserve(relationships.len())
        .map_err(|source| allocation("tag target index", source))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(relationships.len())
        .map_err(|source| allocation("tag source inventory", source))?;
    for relationship in relationships {
        if relationship.is_external() {
            return Err(invalid(format!(
                "tag-list relationship '{}' cannot be external",
                relationship.r_id()
            )));
        }
        let requested_target = relationship.target_partname()?;
        let part = package.get_part(&requested_target)?;
        let target = part.partname().clone();
        if !targets.insert(target.as_str().to_ascii_lowercase()) {
            return Err(invalid(format!(
                "duplicate owner tag-list target '{target}'"
            )));
        }
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
        let (list, conformance) = parse_profiled(part.blob())?;
        output.push(Source {
            relationship_id: relationship.r_id().into(),
            part_name: target,
            conformance,
            list,
        });
    }
    Ok(output)
}

/// Load the optional tag list attached to the selected part-level object.
///
/// For a presentation this follows only direct
/// `p:presentation/p:custDataLst/p:tags`. For slides, layouts, masters, notes,
/// and handouts it follows only direct `p:cSld/p:custDataLst/p:tags`.
/// Shape-level `p:nvPr` anchors and unanchored tag relationships are ignored.
pub fn load(package: &OpcPackage, owner: &PackURI) -> Result<Option<Source>> {
    let owner = package.get_part(owner)?;
    let layout = scan_owner_xml(owner.blob(), owner.content_type())?;
    layout
        .anchor
        .as_ref()
        .map(|anchor| resolve_anchor(owner, package, anchor, layout.conformance))
        .transpose()
}

/// Create or replace the selected part-level object's optional tag list.
///
/// The list is moved into a staged tag part. Adding also stages the owner XML
/// anchor and relationship; replacing preserves the anchor and relationship ID
/// and forks the target only when another package edge shares it. A
/// byte-identical replacement is a signature-preserving no-op. The returned
/// value is the previous list, or `None` when a new attachment was created.
pub fn put(package: &mut OpcPackage, owner: &PackURI, list: List) -> Result<Option<List>> {
    let (owner_name, layout, attached, anchor_uses) = {
        let owner = package.get_part(owner)?;
        let layout = scan_owner_xml(owner.blob(), owner.content_type())?;
        let attached = layout
            .anchor
            .as_ref()
            .map(|anchor| {
                let source = resolve_anchor(owner, package, anchor, layout.conformance)?;
                let relationship = owner.rels().get(source.rel()).ok_or_else(|| {
                    invalid("anchored tag-list relationship disappeared during preflight")
                })?;
                Ok::<_, Error>(Attached {
                    relationship_type: relationship.reltype().into(),
                    source,
                })
            })
            .transpose()?;
        let anchor_uses = layout
            .anchor
            .as_ref()
            .map(|anchor| owner_anchor_uses(owner.blob(), layout.conformance, anchor.id.as_str()))
            .transpose()?
            .unwrap_or(0);
        (owner.partname().clone(), layout, attached, anchor_uses)
    };

    let conformance = attached
        .as_ref()
        .map_or(layout.conformance, |value| value.source.conformance);
    let xml = staged_xml(&list, conformance)?;

    if let Some(attached) = attached {
        if package.get_part(attached.source.part())?.blob() == xml {
            return Ok(Some(attached.source.into_list()));
        }
        let shared_target = has_other_inbound(
            package,
            attached.source.part(),
            &owner_name,
            attached.source.rel(),
        )?;
        let shared_anchor = anchor_uses > 1;
        let fork = shared_target || shared_anchor;
        let part_name = if fork {
            available_part_name(package)?
        } else {
            attached.source.part().clone()
        };
        let target_ref = part_name.relative_ref(owner_name.base_uri());
        validate_relative_target(&owner_name, &target_ref, &part_name)?;
        if fork {
            package.validate_new_part_name(&part_name)?;
        }

        let (relationship_id, owner_xml) = if shared_anchor {
            let owner = package.get_part(&owner_name)?;
            let relationship_id = available_relationship_id(owner)?;
            let owner_xml =
                replace_anchor_relationship_id(owner.blob(), &layout, relationship_id.as_str())?;
            let staged = scan_owner_xml(&owner_xml, owner.content_type())?;
            if staged.conformance != layout.conformance
                || staged.anchor.as_ref().map(|anchor| anchor.id.as_str())
                    != Some(relationship_id.as_str())
            {
                return Err(invalid("staged tag-owner anchor did not round-trip"));
            }
            (relationship_id, Some(owner_xml))
        } else {
            (attached.source.relationship_id.clone(), None)
        };

        {
            let owner = package.get_part_mut(&owner_name)?;
            validate_selected_relationship(
                owner,
                attached.source.rel(),
                &attached.relationship_type,
                attached.source.part(),
            )?;
            if let Some(owner_xml) = owner_xml {
                owner.set_blob(owner_xml);
                owner.rels_mut().add_relationship(
                    attached.relationship_type,
                    target_ref,
                    relationship_id,
                    false,
                );
            } else if shared_target {
                let _ = owner.rels_mut().remove(attached.source.rel());
                owner.rels_mut().add_relationship(
                    attached.relationship_type,
                    target_ref,
                    relationship_id,
                    false,
                );
            }
        }
        package.add_part(Box::new(XmlPart::new(part_name, CONTENT_TYPE.into(), xml)));
        package.unsign();
        return Ok(Some(attached.source.into_list()));
    }

    let (relationship_id, part_name, target_ref, owner_xml) = {
        let owner = package.get_part(&owner_name)?;
        let relationship_id = available_relationship_id(owner)?;
        let part_name = available_part_name(package)?;
        let target_ref = part_name.relative_ref(owner_name.base_uri());
        validate_relative_target(&owner_name, &target_ref, &part_name)?;
        let owner_xml = add_anchor(owner.blob(), &layout, &relationship_id)?;
        let staged = scan_owner_xml(&owner_xml, owner.content_type())?;
        if staged.conformance != layout.conformance
            || staged.anchor.as_ref().map(|anchor| anchor.id.as_str())
                != Some(relationship_id.as_str())
        {
            return Err(invalid("staged tag-owner anchor did not round-trip"));
        }
        (relationship_id, part_name, target_ref, owner_xml)
    };
    package.validate_new_part_name(&part_name)?;
    {
        let owner = package.get_part_mut(&owner_name)?;
        let current = scan_owner_xml(owner.blob(), owner.content_type())?;
        if current.anchor.is_some() || owner.rels().get(&relationship_id).is_some() {
            return Err(invalid("tag-owner graph changed during preflight"));
        }
        owner.set_blob(owner_xml);
        owner.rels_mut().add_relationship(
            layout.conformance.relationship().into(),
            target_ref,
            relationship_id,
            false,
        );
    }
    package.add_part(Box::new(XmlPart::new(part_name, CONTENT_TYPE.into(), xml)));
    package.unsign();
    Ok(None)
}

/// Remove the selected part-level object's optional tag list.
///
/// An absent attachment is an idempotent, signature-preserving `Ok(None)`.
/// Other customer-data children remain byte-for-byte intact; a customer-data
/// container is removed only when the tag anchor was its sole content. The tag
/// part is collected only when no other package edge retains it.
pub fn remove(package: &mut OpcPackage, owner: &PackURI) -> Result<Option<List>> {
    let (owner_name, layout, attached, owner_xml, retain_relationship, orphan) = {
        let owner = package.get_part(owner)?;
        let layout = scan_owner_xml(owner.blob(), owner.content_type())?;
        let Some(anchor) = layout.anchor.as_ref() else {
            return Ok(None);
        };
        let source = resolve_anchor(owner, package, anchor, layout.conformance)?;
        let relationship = owner.rels().get(source.rel()).ok_or_else(|| {
            invalid("anchored tag-list relationship disappeared during preflight")
        })?;
        let relationship_type = relationship.reltype().to_owned();
        let owner_name = owner.partname().clone();
        let retain_relationship =
            owner_anchor_uses(owner.blob(), layout.conformance, anchor.id.as_str())? > 1;
        let owner_xml = remove_anchor(owner.blob(), &layout)?;
        let staged = scan_owner_xml(&owner_xml, owner.content_type())?;
        if staged.conformance != layout.conformance || staged.anchor.is_some() {
            return Err(invalid("staged tag-owner removal did not round-trip"));
        }
        let orphan = !retain_relationship
            && !has_other_inbound(package, source.part(), &owner_name, source.rel())?;
        (
            owner_name,
            layout,
            Attached {
                relationship_type,
                source,
            },
            owner_xml,
            retain_relationship,
            orphan,
        )
    };

    {
        let owner = package.get_part_mut(&owner_name)?;
        let current = scan_owner_xml(owner.blob(), owner.content_type())?;
        if current.anchor.as_ref().map(|anchor| anchor.id.as_str())
            != layout.anchor.as_ref().map(|anchor| anchor.id.as_str())
        {
            return Err(invalid("tag-owner anchor changed during preflight"));
        }
        validate_selected_relationship(
            owner,
            attached.source.rel(),
            &attached.relationship_type,
            attached.source.part(),
        )?;
        owner.set_blob(owner_xml);
        if !retain_relationship {
            let _ = owner.rels_mut().remove(attached.source.rel());
        }
    }
    if orphan {
        let _ = package.remove_part(attached.source.part());
    }
    package.unsign();
    Ok(Some(attached.source.into_list()))
}

struct Attached {
    relationship_type: String,
    source: Source,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum OwnerKind {
    Presentation,
    CommonSlide,
}

#[derive(Debug)]
struct OwnerXml {
    conformance: Conformance,
    insertion: usize,
    container: Option<Container>,
    anchor: Option<Anchor>,
}

#[derive(Debug)]
struct Container {
    span: Range<usize>,
    close_start: usize,
    empty: bool,
    qualified_name: Vec<u8>,
    child_elements: usize,
    other_content: bool,
    preserve_when_empty: bool,
}

#[derive(Debug)]
struct Anchor {
    id: String,
    span: Range<usize>,
}

struct OpenContainer {
    start: usize,
    depth: usize,
    qualified_name: Vec<u8>,
    child_elements: usize,
    other_content: bool,
    preserve_when_empty: bool,
    tags_seen: bool,
}

struct OpenAnchor {
    start: usize,
    depth: usize,
    id: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CommonSlidePhase {
    Start,
    Background,
    Shapes,
    CustomerData,
    Controls,
    Extensions,
}

fn observe_common_slide_child(local: &[u8], phase: &mut CommonSlidePhase) -> Result<()> {
    *phase = match (local, *phase) {
        (b"bg", CommonSlidePhase::Start) => CommonSlidePhase::Background,
        (b"spTree", CommonSlidePhase::Start | CommonSlidePhase::Background) => {
            CommonSlidePhase::Shapes
        },
        (b"custDataLst", CommonSlidePhase::Shapes) => CommonSlidePhase::CustomerData,
        (b"controls", CommonSlidePhase::Shapes | CommonSlidePhase::CustomerData) => {
            CommonSlidePhase::Controls
        },
        (
            b"extLst",
            CommonSlidePhase::Shapes | CommonSlidePhase::CustomerData | CommonSlidePhase::Controls,
        ) => CommonSlidePhase::Extensions,
        (b"spTree", _) => {
            return Err(invalid("direct p:spTree is duplicated or out of order"));
        },
        (b"custDataLst", _) => {
            return Err(invalid(
                "direct p:custDataLst must follow p:spTree and precede later p:cSld children",
            ));
        },
        (b"bg" | b"controls" | b"extLst", _) => {
            return Err(invalid(
                "direct p:cSld children are duplicated or out of order",
            ));
        },
        _ => return Err(invalid("unsupported direct PresentationML p:cSld child")),
    };
    Ok(())
}

fn observe_customer_data_child(is_pml: bool, local: &[u8], tags_seen: &mut bool) -> Result<()> {
    if *tags_seen {
        return Err(invalid("p:tags must be the last p:custDataLst child"));
    }
    if !is_pml {
        return Err(invalid(
            "p:custDataLst contains an unsupported direct child",
        ));
    }
    match local {
        b"custData" => Ok(()),
        b"tags" => {
            *tags_seen = true;
            Ok(())
        },
        _ => Err(invalid(
            "p:custDataLst contains an unsupported direct child",
        )),
    }
}

fn scan_owner_xml(xml: &[u8], content_type: &str) -> Result<OwnerXml> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "tag-owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut kind = None;
    let mut conformance = None;
    let mut root_close = None;
    let mut presentation_insert = None;
    let mut common_slide_depth = None;
    let mut common_slide_seen = false;
    let mut common_slide_phase = CommonSlidePhase::Start;
    let mut after_shape_tree = None;
    let mut shape_tree_depth = None;
    let mut container = None;
    let mut open_container: Option<OpenContainer> = None;
    let mut anchor = None;
    let mut open_anchor: Option<OpenAnchor> = None;

    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let event_conformance = pml(&namespace);
        drop(namespace);
        let end = xml_position(&reader)?;
        match event {
            Event::Start(element) => {
                bump_owner_node(&mut nodes)?;
                if open_anchor.is_some() {
                    return Err(invalid("p:tags cannot contain child elements"));
                }
                if depth == 0 {
                    let found_kind = owner_kind(element.local_name().as_ref(), content_type)?;
                    let found_conformance = event_conformance.ok_or_else(|| {
                        invalid("tag-owner root has an unsupported namespace profile")
                    })?;
                    kind = Some(found_kind);
                    conformance = Some(found_conformance);
                } else {
                    let active_kind = kind.ok_or_else(|| invalid("tag-owner root is missing"))?;
                    let profile =
                        conformance.ok_or_else(|| invalid("tag-owner profile is missing"))?;
                    let is_pml = event_conformance == Some(profile);
                    let local = element.local_name();

                    if active_kind == OwnerKind::CommonSlide
                        && depth == 1
                        && is_pml
                        && local.as_ref() == b"cSld"
                    {
                        if common_slide_seen {
                            return Err(invalid("tag owner has multiple direct p:cSld elements"));
                        }
                        common_slide_seen = true;
                        common_slide_depth = Some(depth + 1);
                    }
                    if active_kind == OwnerKind::Presentation
                        && depth == 1
                        && is_pml
                        && presentation_later(local.as_ref())
                    {
                        presentation_insert.get_or_insert(start);
                    }
                    let target_depth = match active_kind {
                        OwnerKind::Presentation => Some(1),
                        OwnerKind::CommonSlide => common_slide_depth,
                    };
                    if active_kind == OwnerKind::CommonSlide
                        && target_depth == Some(depth)
                        && is_pml
                    {
                        observe_common_slide_child(local.as_ref(), &mut common_slide_phase)?;
                    }
                    if target_depth == Some(depth) && is_pml && local.as_ref() == b"custDataLst" {
                        if container.is_some() || open_container.is_some() {
                            return Err(invalid(
                                "tag owner has multiple direct p:custDataLst elements",
                            ));
                        }
                        open_container = Some(OpenContainer {
                            start,
                            depth: depth + 1,
                            qualified_name: element.name().as_ref().to_vec(),
                            child_elements: 0,
                            other_content: false,
                            preserve_when_empty: has_non_namespace_attrs(&element)?,
                            tags_seen: false,
                        });
                    } else if let Some(current) = open_container.as_mut()
                        && depth == current.depth
                    {
                        observe_customer_data_child(
                            is_pml,
                            local.as_ref(),
                            &mut current.tags_seen,
                        )?;
                        current.child_elements = current.child_elements.saturating_add(1);
                        if is_pml && local.as_ref() == b"tags" {
                            if anchor.is_some() || open_anchor.is_some() {
                                return Err(invalid(
                                    "p:custDataLst contains multiple direct p:tags anchors",
                                ));
                            }
                            open_anchor = Some(OpenAnchor {
                                start,
                                depth: depth + 1,
                                id: anchor_relationship_id(&reader, &element, profile)?,
                            });
                        }
                    }
                    if active_kind == OwnerKind::CommonSlide
                        && common_slide_depth == Some(depth)
                        && is_pml
                        && local.as_ref() == b"spTree"
                    {
                        shape_tree_depth = Some(depth + 1);
                    }
                }
                depth = checked_owner_depth(depth)?;
            },
            Event::Empty(element) => {
                bump_owner_node(&mut nodes)?;
                if open_anchor.is_some() {
                    return Err(invalid("p:tags cannot contain child elements"));
                }
                if depth == 0 {
                    return Err(invalid("tag-owner root cannot be empty"));
                }
                let active_kind = kind.ok_or_else(|| invalid("tag-owner root is missing"))?;
                let profile = conformance.ok_or_else(|| invalid("tag-owner profile is missing"))?;
                let is_pml = event_conformance == Some(profile);
                let local = element.local_name();
                if active_kind == OwnerKind::Presentation
                    && depth == 1
                    && is_pml
                    && presentation_later(local.as_ref())
                {
                    presentation_insert.get_or_insert(start);
                }
                let target_depth = match active_kind {
                    OwnerKind::Presentation => Some(1),
                    OwnerKind::CommonSlide => common_slide_depth,
                };
                if active_kind == OwnerKind::CommonSlide && target_depth == Some(depth) && is_pml {
                    observe_common_slide_child(local.as_ref(), &mut common_slide_phase)?;
                }
                if target_depth == Some(depth) && is_pml && local.as_ref() == b"custDataLst" {
                    if container.is_some() || open_container.is_some() {
                        return Err(invalid(
                            "tag owner has multiple direct p:custDataLst elements",
                        ));
                    }
                    container = Some(Container {
                        span: start..end,
                        close_start: end,
                        empty: true,
                        qualified_name: element.name().as_ref().to_vec(),
                        child_elements: 0,
                        other_content: false,
                        preserve_when_empty: has_non_namespace_attrs(&element)?,
                    });
                } else if let Some(current) = open_container.as_mut()
                    && depth == current.depth
                {
                    observe_customer_data_child(is_pml, local.as_ref(), &mut current.tags_seen)?;
                    current.child_elements = current.child_elements.saturating_add(1);
                    if is_pml && local.as_ref() == b"tags" {
                        if anchor.is_some() || open_anchor.is_some() {
                            return Err(invalid(
                                "p:custDataLst contains multiple direct p:tags anchors",
                            ));
                        }
                        anchor = Some(Anchor {
                            id: anchor_relationship_id(&reader, &element, profile)?,
                            span: start..end,
                        });
                    }
                }
                if active_kind == OwnerKind::CommonSlide
                    && common_slide_depth == Some(depth)
                    && is_pml
                    && local.as_ref() == b"spTree"
                {
                    after_shape_tree = Some(end);
                }
            },
            Event::End(element) => {
                let profile = conformance.ok_or_else(|| invalid("tag-owner profile is missing"))?;
                let is_pml = event_conformance == Some(profile);
                let local = element.local_name();
                if open_anchor
                    .as_ref()
                    .is_some_and(|value| value.depth == depth)
                {
                    if !is_pml || local.as_ref() != b"tags" {
                        return Err(invalid("malformed direct p:tags anchor"));
                    }
                    let value = open_anchor
                        .take()
                        .ok_or_else(|| invalid("tag anchor parser state is inconsistent"))?;
                    anchor = Some(Anchor {
                        id: value.id,
                        span: value.start..end,
                    });
                }
                if open_container
                    .as_ref()
                    .is_some_and(|value| value.depth == depth)
                {
                    if !is_pml || local.as_ref() != b"custDataLst" {
                        return Err(invalid("malformed direct p:custDataLst"));
                    }
                    let value = open_container
                        .take()
                        .ok_or_else(|| invalid("customer-data parser state is inconsistent"))?;
                    container = Some(Container {
                        span: value.start..end,
                        close_start: start,
                        empty: false,
                        qualified_name: value.qualified_name,
                        child_elements: value.child_elements,
                        other_content: value.other_content,
                        preserve_when_empty: value.preserve_when_empty,
                    });
                }
                if shape_tree_depth == Some(depth) {
                    if !is_pml || local.as_ref() != b"spTree" {
                        return Err(invalid("malformed direct p:spTree"));
                    }
                    after_shape_tree = Some(end);
                    shape_tree_depth = None;
                }
                if common_slide_depth == Some(depth) {
                    if !is_pml || local.as_ref() != b"cSld" {
                        return Err(invalid("malformed direct p:cSld"));
                    }
                    if matches!(
                        common_slide_phase,
                        CommonSlidePhase::Start | CommonSlidePhase::Background
                    ) {
                        return Err(invalid("direct p:cSld has no p:spTree"));
                    }
                    common_slide_depth = None;
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("tag-owner XML depth underflow"))?;
            },
            Event::Text(text) => {
                let nonempty = !text.decode().map_err(xml_error)?.trim().is_empty();
                if open_anchor.is_some() && nonempty {
                    return Err(invalid("p:tags cannot contain text"));
                }
                if nonempty
                    && let Some(current) = open_container.as_mut()
                    && depth == current.depth
                {
                    current.other_content = true;
                }
            },
            Event::CData(value) => {
                if open_anchor.is_some() && !value.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("p:tags cannot contain CDATA"));
                }
                if let Some(current) = open_container.as_mut()
                    && depth == current.depth
                {
                    current.other_content = true;
                }
            },
            Event::Comment(_) => {
                if open_anchor.is_none()
                    && let Some(current) = open_container.as_mut()
                    && depth == current.depth
                {
                    current.other_content = true;
                }
            },
            Event::Decl(_) if depth == 0 => {},
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid("tag-owner XML contains forbidden markup"));
            },
            Event::Eof => break,
        }
    }

    if depth != 0 || open_container.is_some() || open_anchor.is_some() {
        return Err(invalid("unterminated tag-owner XML"));
    }
    let kind = kind.ok_or_else(|| invalid("tag-owner root is missing"))?;
    let conformance = conformance.ok_or_else(|| invalid("tag-owner profile is missing"))?;
    let insertion = match kind {
        OwnerKind::Presentation => presentation_insert
            .or(root_close)
            .ok_or_else(|| invalid("presentation root is not closed"))?,
        OwnerKind::CommonSlide => {
            if !common_slide_seen {
                return Err(invalid("tag owner has no direct p:cSld"));
            }
            after_shape_tree.ok_or_else(|| invalid("direct p:cSld has no p:spTree"))?
        },
    };
    Ok(OwnerXml {
        conformance,
        insertion,
        container,
        anchor,
    })
}

fn owner_kind(root: &[u8], content_type: &str) -> Result<OwnerKind> {
    let kind = match root {
        b"presentation"
            if matches!(
                content_type,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                    | "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
                    | "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
                    | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
                    | "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"
                    | "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
                    | "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
            ) =>
        {
            OwnerKind::Presentation
        },
        b"sld"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"sldLayout"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"sldMaster"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"notes"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"notesMaster"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" =>
        {
            OwnerKind::CommonSlide
        },
        b"handoutMaster"
            if content_type
                == "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" =>
        {
            OwnerKind::CommonSlide
        },
        _ => {
            return Err(Error::ContentType {
                expected: "PresentationML programmable-tag owner".into(),
                actual: content_type.into(),
            });
        },
    };
    Ok(kind)
}

fn presentation_later(local: &[u8]) -> bool {
    matches!(
        local,
        b"kinsoku" | b"defaultTextStyle" | b"modifyVerifier" | b"extLst"
    )
}

fn anchor_relationship_id(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: Conformance,
) -> Result<String> {
    let mut relationship_id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() != b"id" || !relationship_namespace(&namespace, conformance) {
            continue;
        }
        if relationship_id.is_some() {
            return Err(invalid("p:tags has duplicate relationship IDs"));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
            return Err(invalid("p:tags has an invalid relationship ID"));
        }
        relationship_id = Some(value);
    }
    relationship_id.ok_or_else(|| invalid("p:tags is missing required r:id"))
}

fn relationship_namespace(namespace: &ResolveResult<'_>, conformance: Conformance) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == conformance.relationship_namespace().as_bytes()
    )
}

fn has_non_namespace_attrs(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name != b"xmlns" && !name.starts_with(b"xmlns:") {
            return Ok(true);
        }
    }
    Ok(false)
}

fn owner_anchor_uses(xml: &[u8], conformance: Conformance, relationship_id: &str) -> Result<usize> {
    let mut reader = NsReader::from_reader(xml);
    let mut nodes = 0usize;
    let mut uses = 0usize;
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let is_pml = pml(&namespace) == Some(conformance);
        drop(namespace);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                bump_owner_node(&mut nodes)?;
                if is_pml && element.local_name().as_ref() == b"tags" {
                    let candidate = anchor_relationship_id(&reader, &element, conformance)?;
                    if candidate == relationship_id {
                        uses = uses.checked_add(1).ok_or(Error::Limit {
                            resource: "tag-owner anchor references",
                            limit: MAX_OWNER_NODES,
                        })?;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(uses)
}

fn resolve_anchor(
    owner: &dyn OpcPart,
    package: &OpcPackage,
    anchor: &Anchor,
    expected: Conformance,
) -> Result<Source> {
    let relationship = owner.rels().get(&anchor.id).ok_or_else(|| {
        invalid(format!(
            "p:tags references missing relationship '{}'",
            anchor.id
        ))
    })?;
    if relationship.reltype() != expected.relationship() {
        return Err(invalid(format!(
            "p:tags relationship '{}' has type '{}' instead of the owner profile's '{}'",
            anchor.id,
            relationship.reltype(),
            expected.relationship(),
        )));
    }
    if relationship.is_external() {
        return Err(invalid("p:tags relationship cannot be external"));
    }
    let requested = relationship.target_partname()?;
    let part = package.get_part(&requested)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(Error::ContentType {
            expected: CONTENT_TYPE.into(),
            actual: part.content_type().into(),
        });
    }
    if !part.rels().is_empty() {
        return Err(invalid(format!(
            "tag-list part '{}' has unexpected relationships",
            part.partname()
        )));
    }
    let (list, conformance) = parse_profiled(part.blob())?;
    if conformance != expected {
        return Err(invalid(
            "tag-list namespace profile does not match its PresentationML owner",
        ));
    }
    Ok(Source {
        relationship_id: anchor.id.clone(),
        part_name: part.partname().clone(),
        conformance,
        list,
    })
}

fn add_anchor(xml: &[u8], layout: &OwnerXml, relationship_id: &str) -> Result<Vec<u8>> {
    if layout.anchor.is_some() {
        return Err(invalid("tag owner already has a direct p:tags anchor"));
    }
    let anchor = format!(
        "<p:tags xmlns:p=\"{}\" xmlns:r=\"{}\" r:id=\"{}\"/>",
        layout.conformance.namespace(),
        layout.conformance.relationship_namespace(),
        relationship_id
    );
    if let Some(container) = &layout.container {
        if !container.empty {
            return insert_xml(xml, container.close_start, anchor.as_bytes());
        }
        let raw = xml
            .get(container.span.clone())
            .ok_or_else(|| invalid("customer-data span is outside owner XML"))?;
        let slash = raw
            .iter()
            .rposition(|byte| *byte == b'/')
            .ok_or_else(|| invalid("empty p:custDataLst has no closing slash"))?;
        let mut replacement = Vec::new();
        replacement.extend_from_slice(&raw[..slash]);
        replacement.extend_from_slice(&raw[slash + 1..]);
        replacement.extend_from_slice(anchor.as_bytes());
        replacement.extend_from_slice(b"</");
        replacement.extend_from_slice(&container.qualified_name);
        replacement.push(b'>');
        return replace_xml(xml, container.span.clone(), &replacement);
    }
    let container = format!(
        "<p:custDataLst xmlns:p=\"{}\">{anchor}</p:custDataLst>",
        layout.conformance.namespace()
    );
    insert_xml(xml, layout.insertion, container.as_bytes())
}

fn replace_anchor_relationship_id(
    xml: &[u8],
    layout: &OwnerXml,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    let anchor = layout
        .anchor
        .as_ref()
        .ok_or_else(|| invalid("tag owner has no direct p:tags anchor"))?;
    let replacement = format!(
        "<p:tags xmlns:p=\"{}\" xmlns:r=\"{}\" r:id=\"{}\"/>",
        layout.conformance.namespace(),
        layout.conformance.relationship_namespace(),
        relationship_id
    );
    replace_xml(xml, anchor.span.clone(), replacement.as_bytes())
}

fn remove_anchor(xml: &[u8], layout: &OwnerXml) -> Result<Vec<u8>> {
    let anchor = layout
        .anchor
        .as_ref()
        .ok_or_else(|| invalid("tag owner has no direct p:tags anchor"))?;
    let container = layout
        .container
        .as_ref()
        .ok_or_else(|| invalid("direct p:tags has no p:custDataLst parent"))?;
    if container.child_elements == 1 && !container.other_content && !container.preserve_when_empty {
        replace_xml(xml, container.span.clone(), &[])
    } else {
        replace_xml(xml, anchor.span.clone(), &[])
    }
}

fn insert_xml(xml: &[u8], offset: usize, value: &[u8]) -> Result<Vec<u8>> {
    replace_xml(xml, offset..offset, value)
}

fn replace_xml(xml: &[u8], range: Range<usize>, value: &[u8]) -> Result<Vec<u8>> {
    let before = xml
        .get(..range.start)
        .ok_or_else(|| invalid("XML patch start is outside owner XML"))?;
    let after = xml
        .get(range.end..)
        .ok_or_else(|| invalid("XML patch end is outside owner XML"))?;
    let len = before
        .len()
        .checked_add(value.len())
        .and_then(|len| len.checked_add(after.len()))
        .ok_or_else(|| invalid("patched tag-owner XML length overflow"))?;
    if len > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "patched tag-owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(len)
        .map_err(|source| allocation("tag-owner XML patch", source))?;
    output.extend_from_slice(before);
    output.extend_from_slice(value);
    output.extend_from_slice(after);
    Ok(output)
}

fn staged_xml(value: &List, conformance: Conformance) -> Result<Vec<u8>> {
    let xml = write(value, conformance)?;
    let (staged, staged_conformance) = parse_profiled(&xml)?;
    if staged_conformance != conformance || &staged != value {
        return Err(invalid("staged tag-list XML did not round-trip"));
    }
    Ok(xml)
}

fn available_relationship_id(owner: &dyn OpcPart) -> Result<String> {
    if owner.rels().len() >= MAX_SOURCE_RELATIONSHIPS {
        return Err(Error::Limit {
            resource: "tag-owner relationships",
            limit: MAX_SOURCE_RELATIONSHIPS,
        });
    }
    for number in 1..=MAX_SOURCE_RELATIONSHIPS {
        let candidate = format!("rId{number}");
        if owner.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        resource: "tag-list relationship-ID allocation attempts",
        limit: MAX_SOURCE_RELATIONSHIPS,
    })
}

fn available_part_name(package: &OpcPackage) -> Result<PackURI> {
    let existing = sorted_part_names(package)?;
    for number in 1..=PART_NAME_ATTEMPTS {
        let path = format!("/ppt/tags/tag{number}.xml");
        let candidate = PackURI::new(&path)
            .map_err(|error| invalid(format!("invalid generated tag-list part name: {error}")))?;
        if !part_name_conflicts(&existing, &path.to_ascii_lowercase()) {
            package.validate_new_part_name(&candidate)?;
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        resource: "tag-list part-name allocation attempts",
        limit: PART_NAME_ATTEMPTS,
    })
}

fn sorted_part_names(package: &OpcPackage) -> Result<Vec<String>> {
    if package.part_count() > MAX_GRAPH_PARTS {
        return Err(Error::Limit {
            resource: "tag package part-name scan",
            limit: MAX_GRAPH_PARTS,
        });
    }
    let mut names = Vec::new();
    names
        .try_reserve_exact(package.part_count())
        .map_err(|source| allocation("tag package part-name index", source))?;
    let mut bytes = 0usize;
    for part in package.iter_parts() {
        bytes = bytes
            .checked_add(part.partname().as_str().len())
            .ok_or(Error::Limit {
                resource: "tag package part-name bytes",
                limit: MAX_PART_NAME_BYTES,
            })?;
        if bytes > MAX_PART_NAME_BYTES {
            return Err(Error::Limit {
                resource: "tag package part-name bytes",
                limit: MAX_PART_NAME_BYTES,
            });
        }
        names.push(part.partname().as_str().to_ascii_lowercase());
    }
    names.sort_unstable();
    names.dedup();
    Ok(names)
}

fn part_name_conflicts(existing: &[String], candidate: &str) -> bool {
    if existing
        .binary_search_by(|name| name.as_str().cmp(candidate))
        .is_ok()
    {
        return true;
    }
    for (index, _) in candidate.match_indices('/').skip(1) {
        if existing
            .binary_search_by(|name| name.as_str().cmp(&candidate[..index]))
            .is_ok()
        {
            return true;
        }
    }
    let descendant = format!("{candidate}/");
    let position = existing.partition_point(|name| name.as_str() < descendant.as_str());
    existing
        .get(position)
        .is_some_and(|name| name.starts_with(&descendant))
}

fn validate_relative_target(source: &PackURI, reference: &str, target: &PackURI) -> Result<()> {
    let resolved = PackURI::from_rel_ref(source.base_uri(), reference)
        .map_err(|error| invalid(format!("invalid generated tag-list target: {error}")))?;
    if resolved.is_equivalent_to(target) {
        Ok(())
    } else {
        Err(invalid("generated tag-list target resolves incorrectly"))
    }
}

fn validate_selected_relationship(
    owner: &dyn OpcPart,
    relationship_id: &str,
    relationship_type: &str,
    target: &PackURI,
) -> Result<()> {
    let relationship = owner.rels().get(relationship_id).ok_or_else(|| {
        invalid(format!(
            "tag-list relationship '{relationship_id}' is missing"
        ))
    })?;
    if relationship.is_external()
        || relationship.reltype() != relationship_type
        || !relationship.target_partname()?.is_equivalent_to(target)
    {
        return Err(invalid(
            "anchored tag-list relationship changed during preflight",
        ));
    }
    Ok(())
}

fn has_other_inbound(
    package: &OpcPackage,
    target: &PackURI,
    selected_owner: &PackURI,
    selected_relationship: &str,
) -> Result<bool> {
    if package.part_count() > MAX_GRAPH_PARTS {
        return Err(Error::Limit {
            resource: "tag package graph parts",
            limit: MAX_GRAPH_PARTS,
        });
    }
    let mut scanned = 0usize;
    for relationship in package.rels().iter() {
        bump_graph_link(&mut scanned)?;
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            return Ok(true);
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            bump_graph_link(&mut scanned)?;
            if source.partname().is_equivalent_to(selected_owner)
                && relationship.r_id() == selected_relationship
            {
                continue;
            }
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn bump_graph_link(scanned: &mut usize) -> Result<()> {
    *scanned = scanned.checked_add(1).ok_or(Error::Limit {
        resource: "tag package graph relationships",
        limit: MAX_GRAPH_LINKS,
    })?;
    if *scanned > MAX_GRAPH_LINKS {
        Err(Error::Limit {
            resource: "tag package graph relationships",
            limit: MAX_GRAPH_LINKS,
        })
    } else {
        Ok(())
    }
}

fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| invalid("tag-owner XML offset overflow"))
}

fn bump_owner_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "tag-owner XML nodes",
        limit: MAX_OWNER_NODES,
    })?;
    if *nodes > MAX_OWNER_NODES {
        Err(Error::Limit {
            resource: "tag-owner XML nodes",
            limit: MAX_OWNER_NODES,
        })
    } else {
        Ok(())
    }
}

fn checked_owner_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("tag-owner XML depth overflow"))?;
    if depth > MAX_OWNER_DEPTH {
        Err(Error::Limit {
            resource: "tag-owner XML depth",
            limit: MAX_OWNER_DEPTH,
        })
    } else {
        Ok(depth)
    }
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
    let mut seen = HashMap::<String, usize>::new();
    seen.try_reserve(value.tags.len())
        .map_err(|source| allocation("tag-name validation index", source))?;
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

fn pml(namespace: &ResolveResult<'_>) -> Option<Conformance> {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == PML => Some(Conformance::Transitional),
        ResolveResult::Bound(value) if value.as_ref() == STRICT => Some(Conformance::Strict),
        _ => None,
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
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

    fn owner_part(
        part_name: &str,
        root: &str,
        content_type: &str,
        conformance: Conformance,
    ) -> XmlPart {
        let body = if root == "presentation" {
            String::new()
        } else {
            "<p:cSld><p:spTree/></p:cSld>".to_owned()
        };
        XmlPart::new(
            PackURI::new(part_name).unwrap(),
            content_type.into(),
            format!(
                r#"<p:{root} xmlns:p="{}" xmlns:r="{}">{body}</p:{root}>"#,
                conformance.namespace(),
                conformance.relationship_namespace(),
            )
            .into_bytes(),
        )
    }

    fn package_with_slide(conformance: Conformance) -> (OpcPackage, PackURI) {
        let name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(owner_part(
            name.as_str(),
            "sld",
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            conformance,
        )));
        (package, name)
    }

    fn list(name: &str, value: &str) -> List {
        let mut list = List::new();
        list.add(Tag::new(name, value).unwrap()).unwrap();
        list
    }

    fn mark_signed(package: &mut OpcPackage) {
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());
    }

    #[test]
    fn anchored_crud_is_strict_profiled_and_signature_safe() {
        use std::sync::Arc;

        let (mut package, owner) = package_with_slide(Conformance::Strict);
        assert_eq!(load(&package, &owner).unwrap(), None);
        assert_eq!(
            put(&mut package, &owner, list("Owner", "Alice")).unwrap(),
            None
        );
        let created = load(&package, &owner).unwrap().unwrap();
        assert_eq!(created.conformance(), Conformance::Strict);
        assert_eq!(created.rel(), "rId1");
        assert_eq!(created.part().as_str(), "/ppt/tags/tag1.xml");
        let relationship = package
            .get_part(&owner)
            .unwrap()
            .rels()
            .get(created.rel())
            .unwrap();
        assert_eq!(relationship.reltype(), STRICT_TAG_REL);
        assert!(
            std::str::from_utf8(package.get_part(created.part()).unwrap().blob())
                .unwrap()
                .contains(STRICT_TEXT)
        );
        let owner_xml = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
        assert!(owner_xml.contains("<p:custDataLst"));
        assert!(owner_xml.contains("<p:tags"));

        let owner_before = package.get_part(&owner).unwrap().blob_arc();
        let part_before = package.get_part(created.part()).unwrap().blob_arc();
        mark_signed(&mut package);
        let old = put(&mut package, &owner, list("Owner", "Alice"))
            .unwrap()
            .unwrap();
        assert_eq!(old.get("owner").unwrap().value(), "Alice");
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &owner_before,
            &package.get_part(&owner).unwrap().blob_arc()
        ));
        assert!(Arc::ptr_eq(
            &part_before,
            &package.get_part(created.part()).unwrap().blob_arc()
        ));

        let malformed = format!(
            r#"<p:tagLst xmlns:p="{STRICT_TEXT}"><p:tag name="Owner" val="one"/><p:tag name="OWNER" val="two"/></p:tagLst>"#
        );
        let malformed = parse(malformed.as_bytes()).unwrap();
        assert!(matches!(
            put(&mut package, &owner, malformed),
            Err(Error::DuplicateName { .. })
        ));
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &part_before,
            &package.get_part(created.part()).unwrap().blob_arc()
        ));

        let old = put(&mut package, &owner, list("Reviewer", "Bob"))
            .unwrap()
            .unwrap();
        assert_eq!(old.get("owner").unwrap().value(), "Alice");
        assert!(!package.is_signed());
        assert_eq!(
            load(&package, &owner)
                .unwrap()
                .unwrap()
                .list()
                .get("reviewer")
                .unwrap()
                .value(),
            "Bob"
        );

        mark_signed(&mut package);
        let removed = remove(&mut package, &owner).unwrap().unwrap();
        assert_eq!(removed.get("reviewer").unwrap().value(), "Bob");
        assert!(!package.is_signed());
        assert!(package.get_part(created.part()).is_err());
        assert_eq!(load(&package, &owner).unwrap(), None);

        let after_remove = package.get_part(&owner).unwrap().blob_arc();
        mark_signed(&mut package);
        assert_eq!(remove(&mut package, &owner).unwrap(), None);
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &after_remove,
            &package.get_part(&owner).unwrap().blob_arc()
        ));
    }

    #[test]
    fn customer_data_and_schema_order_are_preserved() {
        let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let original = format!(
            r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}"><p:cSld><p:spTree/><p:custDataLst keep="yes"><p:custData r:id="rIdData"/></p:custDataLst><p:controls/></p:cSld></p:sld>"#
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            owner.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            original.as_bytes().to_vec(),
        )));

        assert_eq!(
            put(&mut package, &owner, list("Owner", "Alice")).unwrap(),
            None
        );
        let updated = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
        let customer = updated.find("<p:custData ").unwrap();
        let tags = updated.find("<p:tags ").unwrap();
        let controls = updated.find("<p:controls").unwrap();
        assert!(customer < tags && tags < controls);

        assert!(remove(&mut package, &owner).unwrap().is_some());
        assert_eq!(
            package.get_part(&owner).unwrap().blob(),
            original.as_bytes()
        );

        let empty = format!(
            r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}"><p:cSld><p:spTree/><p:custDataLst keep="yes"/><p:controls/></p:cSld></p:sld>"#
        );
        package
            .get_part_mut(&owner)
            .unwrap()
            .set_blob(empty.into_bytes());
        assert_eq!(put(&mut package, &owner, List::new()).unwrap(), None);
        assert_eq!(remove(&mut package, &owner).unwrap(), Some(List::new()));
        let restored = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
        assert!(restored.contains("<p:custDataLst keep=\"yes\"></p:custDataLst>"));
        assert!(restored.contains("<p:controls"));
    }

    #[test]
    fn malformed_owner_order_is_rejected_before_mutation() {
        use std::sync::Arc;

        let cases = [
            "<p:cSld><p:custDataLst/><p:spTree/></p:cSld>",
            "<p:cSld><p:spTree/><p:spTree/></p:cSld>",
            r#"<p:cSld><p:spTree/><p:custDataLst><p:tags r:id="rId1"/><p:custData r:id="rIdData"/></p:custDataLst></p:cSld>"#,
        ];
        for body in cases {
            let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
            let mut package = OpcPackage::new();
            package.add_part(Box::new(XmlPart::new(
                owner.clone(),
                "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
                format!(r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}">{body}</p:sld>"#,)
                    .into_bytes(),
            )));
            let owner_before = package.get_part(&owner).unwrap().blob_arc();
            mark_signed(&mut package);

            assert!(matches!(load(&package, &owner), Err(Error::Invalid(_))));
            assert!(matches!(
                put(&mut package, &owner, list("Owner", "Alice")),
                Err(Error::Invalid(_))
            ));
            assert!(matches!(
                remove(&mut package, &owner),
                Err(Error::Invalid(_))
            ));
            assert!(package.is_signed());
            assert!(Arc::ptr_eq(
                &owner_before,
                &package.get_part(&owner).unwrap().blob_arc()
            ));
        }
    }

    #[test]
    fn package_shared_target_forks_and_collects_only_orphans() {
        let (mut package, first_owner) = package_with_slide(Conformance::Transitional);
        assert_eq!(
            put(&mut package, &first_owner, list("Owner", "Alice")).unwrap(),
            None
        );
        let original = load(&package, &first_owner).unwrap().unwrap();
        let original_part = original.part().clone();
        let original_bytes = package.get_part(&original_part).unwrap().blob_arc();

        let second_owner = PackURI::new("/ppt/slides/slide2.xml").unwrap();
        let second_xml = format!(
            r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}"><p:cSld><p:spTree/><p:custDataLst><p:tags r:id="rIdShared"/></p:custDataLst></p:cSld></p:sld>"#
        );
        package.add_part(Box::new(XmlPart::new(
            second_owner.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            second_xml.into_bytes(),
        )));
        package
            .get_part_mut(&second_owner)
            .unwrap()
            .rels_mut()
            .add_relationship(
                TAG_REL.into(),
                original_part.relative_ref(second_owner.base_uri()),
                "rIdShared".into(),
                false,
            );

        assert_eq!(
            load(&package, &second_owner)
                .unwrap()
                .unwrap()
                .list()
                .get("owner")
                .unwrap()
                .value(),
            "Alice"
        );
        let old = put(&mut package, &first_owner, list("Reviewer", "Bob"))
            .unwrap()
            .unwrap();
        assert_eq!(old.get("owner").unwrap().value(), "Alice");
        let first = load(&package, &first_owner).unwrap().unwrap();
        let second = load(&package, &second_owner).unwrap().unwrap();
        assert_ne!(first.part(), &original_part);
        assert_eq!(second.part(), &original_part);
        assert_eq!(first.list().get("reviewer").unwrap().value(), "Bob");
        assert_eq!(second.list().get("owner").unwrap().value(), "Alice");
        assert!(std::sync::Arc::ptr_eq(
            &original_bytes,
            &package.get_part(&original_part).unwrap().blob_arc()
        ));

        let fork = first.part().clone();
        assert!(remove(&mut package, &first_owner).unwrap().is_some());
        assert!(package.get_part(&fork).is_err());
        assert!(package.get_part(&original_part).is_ok());
        assert!(remove(&mut package, &second_owner).unwrap().is_some());
        assert!(package.get_part(&original_part).is_err());
    }

    #[test]
    fn same_owner_reused_anchor_forks_relationship_and_part() {
        let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let original_part = PackURI::new("/ppt/tags/tag1.xml").unwrap();
        let owner_xml = format!(
            r#"<p:sld xmlns:p="{PML_TEXT}" xmlns:r="{REL_TEXT}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="1" name="Shape"/><p:cNvSpPr/><p:nvPr><p:custDataLst><p:tags r:id="rIdShared"/></p:custDataLst></p:nvPr></p:nvSpPr></p:sp></p:spTree><p:custDataLst><p:tags r:id="rIdShared"/></p:custDataLst></p:cSld></p:sld>"#
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            owner.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            owner_xml.into_bytes(),
        )));
        package.add_part(Box::new(XmlPart::new(
            original_part.clone(),
            CONTENT_TYPE.into(),
            write(&list("Owner", "Alice"), Conformance::Transitional).unwrap(),
        )));
        package
            .get_part_mut(&owner)
            .unwrap()
            .rels_mut()
            .add_relationship(
                TAG_REL.into(),
                original_part.relative_ref(owner.base_uri()),
                "rIdShared".into(),
                false,
            );

        let old = put(&mut package, &owner, list("Reviewer", "Bob"))
            .unwrap()
            .unwrap();
        assert_eq!(old.get("owner").unwrap().value(), "Alice");
        let direct = load(&package, &owner).unwrap().unwrap();
        assert_ne!(direct.rel(), "rIdShared");
        assert_ne!(direct.part(), &original_part);
        assert_eq!(direct.list().get("reviewer").unwrap().value(), "Bob");
        assert_eq!(
            discover(package.get_part(&owner).unwrap(), &package)
                .unwrap()
                .len(),
            2
        );
        let updated = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
        assert_eq!(updated.matches("rIdShared").count(), 1);
        assert!(package.get_part(&original_part).is_ok());

        let fork = direct.part().clone();
        assert!(remove(&mut package, &owner).unwrap().is_some());
        assert_eq!(load(&package, &owner).unwrap(), None);
        assert!(package.get_part(&fork).is_err());
        assert!(package.get_part(&original_part).is_ok());
        assert!(
            package
                .get_part(&owner)
                .unwrap()
                .rels()
                .get("rIdShared")
                .is_some()
        );
        assert_eq!(
            discover(package.get_part(&owner).unwrap(), &package)
                .unwrap()
                .len(),
            1
        );
    }

    #[test]
    fn mixed_profile_preflight_is_atomic() {
        use std::sync::Arc;

        let (mut package, owner) = package_with_slide(Conformance::Strict);
        assert_eq!(
            put(&mut package, &owner, list("Owner", "Alice")).unwrap(),
            None
        );
        let part_name = load(&package, &owner).unwrap().unwrap().part().clone();
        package
            .get_part_mut(&part_name)
            .unwrap()
            .set_blob(write(&list("Owner", "Alice"), Conformance::Transitional).unwrap());
        let owner_before = package.get_part(&owner).unwrap().blob_arc();
        let part_before = package.get_part(&part_name).unwrap().blob_arc();
        mark_signed(&mut package);

        for result in [
            load(&package, &owner).map(|_| ()),
            put(&mut package, &owner, list("Reviewer", "Bob")).map(|_| ()),
            remove(&mut package, &owner).map(|_| ()),
        ] {
            assert!(matches!(
                result,
                Err(Error::Invalid(message)) if message.contains("namespace profile")
            ));
        }
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &owner_before,
            &package.get_part(&owner).unwrap().blob_arc()
        ));
        assert!(Arc::ptr_eq(
            &part_before,
            &package.get_part(&part_name).unwrap().blob_arc()
        ));
    }

    #[test]
    fn mixed_owner_relationship_profiles_are_rejected_atomically() {
        use std::sync::Arc;

        for (relationship_namespace, relationship_type) in
            [(REL_TEXT, STRICT_TAG_REL), (STRICT_REL_TEXT, TAG_REL)]
        {
            let owner = PackURI::new("/ppt/slides/slide1.xml").unwrap();
            let part_name = PackURI::new("/ppt/tags/tag1.xml").unwrap();
            let owner_xml = format!(
                r#"<p:sld xmlns:p="{STRICT_TEXT}" xmlns:r="{relationship_namespace}"><p:cSld><p:spTree/><p:custDataLst><p:tags r:id="rId1"/></p:custDataLst></p:cSld></p:sld>"#,
            );
            let mut package = OpcPackage::new();
            package.add_part(Box::new(XmlPart::new(
                owner.clone(),
                "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
                owner_xml.into_bytes(),
            )));
            package.add_part(Box::new(XmlPart::new(
                part_name.clone(),
                CONTENT_TYPE.into(),
                write(&list("Owner", "Alice"), Conformance::Strict).unwrap(),
            )));
            package
                .get_part_mut(&owner)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    relationship_type.into(),
                    part_name.relative_ref(owner.base_uri()),
                    "rId1".into(),
                    false,
                );
            let owner_before = package.get_part(&owner).unwrap().blob_arc();
            let part_before = package.get_part(&part_name).unwrap().blob_arc();
            mark_signed(&mut package);

            assert!(matches!(load(&package, &owner), Err(Error::Invalid(_))));
            assert!(matches!(
                put(&mut package, &owner, list("Reviewer", "Bob")),
                Err(Error::Invalid(_))
            ));
            assert!(matches!(
                remove(&mut package, &owner),
                Err(Error::Invalid(_))
            ));
            assert!(package.is_signed());
            assert!(Arc::ptr_eq(
                &owner_before,
                &package.get_part(&owner).unwrap().blob_arc()
            ));
            assert!(Arc::ptr_eq(
                &part_before,
                &package.get_part(&part_name).unwrap().blob_arc()
            ));
        }
    }

    #[test]
    fn creation_supports_all_presentationml_tag_owners_and_empty_lists() {
        let owners = [
            (
                "/ppt/presentation.xml",
                "presentation",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            ),
            (
                "/ppt/slides/slide1.xml",
                "sld",
                "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            ),
            (
                "/ppt/slideLayouts/slideLayout1.xml",
                "sldLayout",
                "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
            ),
            (
                "/ppt/slideMasters/slideMaster1.xml",
                "sldMaster",
                "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
            ),
            (
                "/ppt/notesSlides/notesSlide1.xml",
                "notes",
                "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
            ),
            (
                "/ppt/notesMasters/notesMaster1.xml",
                "notesMaster",
                "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
            ),
            (
                "/ppt/handoutMasters/handoutMaster1.xml",
                "handoutMaster",
                "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml",
            ),
        ];
        let mut package = OpcPackage::new();
        for (part_name, root, content_type) in owners {
            let owner = PackURI::new(part_name).unwrap();
            package.add_part(Box::new(owner_part(
                part_name,
                root,
                content_type,
                Conformance::Transitional,
            )));
            assert_eq!(put(&mut package, &owner, List::new()).unwrap(), None);
            let source = load(&package, &owner).unwrap().unwrap();
            assert!(source.list().is_empty());
            assert_eq!(source.conformance(), Conformance::Transitional);
            assert_eq!(
                discover(package.get_part(&owner).unwrap(), &package)
                    .unwrap()
                    .len(),
                1
            );
            assert_eq!(remove(&mut package, &owner).unwrap(), Some(List::new()));
            assert_eq!(load(&package, &owner).unwrap(), None);
        }
    }

    #[test]
    fn part_allocation_avoids_ascii_case_and_derived_name_collisions() {
        use litchi_opc::BlobPart;

        let (mut package, owner) = package_with_slide(Conformance::Transitional);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/PPT/TAGS/TAG1.XML").unwrap(),
            "application/octet-stream".into(),
            Vec::new(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/tags/tag2.xml/child").unwrap(),
            "application/octet-stream".into(),
            Vec::new(),
        )));

        assert_eq!(put(&mut package, &owner, List::new()).unwrap(), None);
        let source = load(&package, &owner).unwrap().unwrap();
        assert_eq!(source.part().as_str(), "/ppt/tags/tag3.xml");
    }
}
