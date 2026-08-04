use super::*;
use crate::{Error, Result};
use caseless::Caseless;
use litchi_opc::PackURI;
use std::collections::{HashMap, HashSet};
use unicode_normalization::UnicodeNormalization;

/// Namespace profile used when a detached list is serialized.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    /// ECMA-376 transitional namespace profile.
    Transitional,
    /// ISO/IEC 29500 strict namespace profile.
    Strict,
}

impl Conformance {
    pub(crate) fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => PML_TEXT,
            Self::Strict => STRICT_TEXT,
        }
    }

    pub(crate) fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => TAG_REL,
            Self::Strict => STRICT_TAG_REL,
        }
    }

    pub(crate) fn relationship_namespace(self) -> &'static str {
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
        pub(in crate::tag) qualified_name: String,
        pub(in crate::tag) value: String,
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
    pub(in crate::tag) name: String,
    pub(in crate::tag) value: String,
    pub(in crate::tag) namespaces: Vec<raw::Attr>,
    pub(in crate::tag) attrs: Vec<raw::Attr>,
    pub(in crate::tag) wire_len: usize,
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
    pub(in crate::tag) tags: Vec<Tag>,
    pub(in crate::tag) namespaces: Vec<raw::Attr>,
    pub(in crate::tag) attrs: Vec<raw::Attr>,
    pub(in crate::tag) wire_len: usize,
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
    pub(in crate::tag) relationship_id: String,
    pub(in crate::tag) part_name: PackURI,
    pub(in crate::tag) conformance: Conformance,
    pub(in crate::tag) list: List,
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

pub(crate) fn unique_offset(tags: &[Tag], name: &str) -> Result<usize> {
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

pub(crate) fn same_name(left: &str, right: &str) -> bool {
    name_key(left) == name_key(right)
}

pub(crate) fn name_key(value: &str) -> String {
    value.chars().nfd().default_case_fold().nfd().collect()
}

pub(crate) fn escaped_len(value: &str) -> Result<usize> {
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

pub(crate) fn attr_wire_len(attr: &raw::Attr) -> Result<usize> {
    attr_wire_len_parts(attr.qualified_name(), attr.value())
}

pub(crate) fn attr_wire_len_parts(name: &str, value: &str) -> Result<usize> {
    checked_wire_add(
        name.len()
            .checked_add(4)
            .ok_or_else(|| invalid("tag attribute length overflow"))?,
        escaped_len(value)?,
    )
}

pub(crate) fn tag_wire_len_parts(
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

pub(crate) fn list_wire_len_parts(
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

pub(crate) fn checked_wire_add(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid("tag-list wire length overflow"))
}

pub(crate) fn checked_wire_sub(left: usize, right: usize) -> Result<usize> {
    left.checked_sub(right)
        .ok_or_else(|| invalid("tag-list wire length invariant underflow"))
}

pub(crate) fn replace_wire_component(current: usize, old: usize, new: usize) -> Result<usize> {
    checked_wire_add(checked_wire_sub(current, old)?, new)
}

pub(crate) fn ensure_tag_budget(wire_len: usize) -> Result<()> {
    ensure_list_budget(checked_wire_add(NONEMPTY_WIRE_BYTES, wire_len)?)
}

pub(crate) fn ensure_list_budget(wire_len: usize) -> Result<()> {
    if wire_len <= MAX_PART_BYTES {
        Ok(())
    } else {
        Err(Error::Limit {
            resource: "encoded tag-list bytes",
            limit: MAX_PART_BYTES,
        })
    }
}

pub(crate) fn validate_structure(value: &List) -> Result<()> {
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

pub(crate) fn validate_unique_names(value: &List) -> Result<()> {
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

pub(crate) fn validate_tag(tag: &Tag) -> Result<()> {
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

pub(crate) fn validate_tag_context(tag: &Tag, root_prefixes: &[String]) -> Result<()> {
    validate_element_attrs(&tag.namespaces, &tag.attrs, &["name", "val"])?;
    let prefixes = bound_prefixes(&tag.namespaces, root_prefixes)?;
    validate_bound_attrs(&tag.attrs, &prefixes)
}

pub(crate) fn validate_element_attrs(
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

pub(crate) fn validate_bound_attrs(attrs: &[raw::Attr], prefixes: &[String]) -> Result<()> {
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

pub(crate) fn check_namespace(
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

pub(crate) fn check_extension(
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
