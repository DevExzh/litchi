//! Package-independent Designer design-element values.

use crate::{Error, Result};

const MAX_UNKNOWN_EXTENSIONS: usize = 1_024;
const MAX_UNKNOWN_EXTENSION_BYTES: usize = 1 << 20;
const MAX_UNKNOWN_BYTES: usize = 8 << 20;

const DEFAULT_TAG_COUNT: usize = 1_024;
const DEFAULT_STRING_BYTES: usize = 64 * 1024;
const DEFAULT_AGGREGATE_BYTES: usize = 1 << 20;
const DEFAULT_XML_BYTES: usize = 8 << 20;
const DEFAULT_XML_NODES: usize = 65_536;
const DEFAULT_XML_DEPTH: usize = 128;
const DEFAULT_ATTRIBUTE_BYTES: usize = 1 << 20;

/// Resource bounds for detached Designer drawing properties.
///
/// The aggregate bound counts UTF-8 bytes in tag names and values. Empty
/// strings are valid and do not consume that budget.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    tag_count: usize,
    string_bytes: usize,
    aggregate_bytes: usize,
    xml_bytes: usize,
    xml_nodes: usize,
    xml_depth: usize,
    attribute_bytes: usize,
}

impl Limits {
    /// Construct the safe default Designer bounds.
    #[inline]
    pub const fn new() -> Self {
        Self {
            tag_count: DEFAULT_TAG_COUNT,
            string_bytes: DEFAULT_STRING_BYTES,
            aggregate_bytes: DEFAULT_AGGREGATE_BYTES,
            xml_bytes: DEFAULT_XML_BYTES,
            xml_nodes: DEFAULT_XML_NODES,
            xml_depth: DEFAULT_XML_DEPTH,
            attribute_bytes: DEFAULT_ATTRIBUTE_BYTES,
        }
    }

    /// Return the maximum number of tags in one tag collection.
    #[inline]
    pub const fn tag_count(self) -> usize {
        self.tag_count
    }

    /// Return the maximum UTF-8 byte length of one tag name or value.
    #[inline]
    pub const fn string_bytes(self) -> usize {
        self.string_bytes
    }

    /// Return the maximum combined UTF-8 byte length of all tag strings.
    #[inline]
    pub const fn aggregate_bytes(self) -> usize {
        self.aggregate_bytes
    }

    /// Maximum bytes in one Designer payload, before or after authoring.
    #[inline]
    pub const fn xml_bytes(self) -> usize {
        self.xml_bytes
    }

    /// Maximum XML events in one Designer payload.
    #[inline]
    pub const fn xml_nodes(self) -> usize {
        self.xml_nodes
    }

    /// Maximum element nesting depth in one Designer payload.
    #[inline]
    pub const fn xml_depth(self) -> usize {
        self.xml_depth
    }

    /// Maximum raw or decoded bytes in one Designer attribute.
    #[inline]
    pub const fn attribute_bytes(self) -> usize {
        self.attribute_bytes
    }

    /// Set the maximum number of tags in one collection.
    #[must_use]
    #[inline]
    pub const fn with_tag_count(mut self, value: usize) -> Self {
        self.tag_count = value;
        self
    }

    /// Set the maximum UTF-8 byte length of one tag name or value.
    #[must_use]
    #[inline]
    pub const fn with_string_bytes(mut self, value: usize) -> Self {
        self.string_bytes = value;
        self
    }

    /// Set the maximum combined UTF-8 byte length of all tag strings.
    #[must_use]
    #[inline]
    pub const fn with_aggregate_bytes(mut self, value: usize) -> Self {
        self.aggregate_bytes = value;
        self
    }

    /// Set the input/output payload-byte limit.
    #[must_use]
    #[inline]
    pub const fn with_xml_bytes(mut self, value: usize) -> Self {
        self.xml_bytes = value;
        self
    }

    /// Set the XML event-count limit.
    #[must_use]
    #[inline]
    pub const fn with_xml_nodes(mut self, value: usize) -> Self {
        self.xml_nodes = value;
        self
    }

    /// Set the XML nesting-depth limit.
    #[must_use]
    #[inline]
    pub const fn with_xml_depth(mut self, value: usize) -> Self {
        self.xml_depth = value;
        self
    }

    /// Set the raw/decoded per-attribute byte limit.
    #[must_use]
    #[inline]
    pub const fn with_attribute_bytes(mut self, value: usize) -> Self {
        self.attribute_bytes = value;
        self
    }
}

impl Default for Limits {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// One inert Designer tag.
///
/// Names and values are retained as XML 1.0 character data. Their meaning is
/// deliberately not interpreted by this package-independent layer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Tag {
    name: String,
    value: String,
}

impl Tag {
    /// Construct a tag under the default resource bounds.
    pub fn new(name: impl Into<String>, value: impl Into<String>) -> Result<Self> {
        Self::new_with_limits(name, value, Limits::default())
    }

    /// Construct a tag under caller-supplied resource bounds.
    pub fn new_with_limits(
        name: impl Into<String>,
        value: impl Into<String>,
        limits: Limits,
    ) -> Result<Self> {
        let name = name.into();
        let value = value.into();
        super::validation::validate_tag_text(&name, limits, "designer tag name")?;
        super::validation::validate_tag_text(&value, limits, "designer tag value")?;
        Ok(Self { name, value })
    }

    /// Return the inert name exactly as supplied.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the inert value exactly as supplied.
    #[inline]
    pub fn value(&self) -> &str {
        &self.value
    }

    fn bytes(&self) -> Result<usize> {
        self.name
            .len()
            .checked_add(self.value.len())
            .ok_or(Error::Limit {
                resource: "designer tag aggregate bytes",
                limit: usize::MAX,
            })
    }

    fn validate(&self, limits: Limits) -> Result<usize> {
        super::validation::validate_tag_text(&self.name, limits, "designer tag name")?;
        super::validation::validate_tag_text(&self.value, limits, "designer tag value")?;
        self.bytes()
    }
}

/// An ordered Designer tag collection.
///
/// Empty collections and duplicate name/value pairs are retained exactly as
/// supplied. The optional collection on [`DrawingProperties`] distinguishes
/// an absent tag list from a present empty one.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Tags {
    entries: Vec<Tag>,
    bytes: usize,
}

impl Tags {
    /// Construct a present, empty tag collection.
    #[inline]
    pub const fn new() -> Self {
        Self {
            entries: Vec::new(),
            bytes: 0,
        }
    }

    /// Construct an ordered tag collection under caller-supplied bounds.
    pub fn from_iter_with_limits(
        entries: impl IntoIterator<Item = Tag>,
        limits: Limits,
    ) -> Result<Self> {
        let mut tags = Self::new();
        for entry in entries {
            tags.push_with_limits(entry, limits)?;
        }
        Ok(tags)
    }

    /// Return the tags in source order.
    #[inline]
    pub fn as_slice(&self) -> &[Tag] {
        &self.entries
    }

    /// Iterate over tags in source order.
    #[inline]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Tag> {
        self.entries.iter()
    }

    /// Return the number of tags.
    #[inline]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Report whether this present collection is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Return the combined UTF-8 byte length of stored names and values.
    #[inline]
    pub const fn aggregate_bytes(&self) -> usize {
        self.bytes
    }

    /// Append a tag under the default resource bounds.
    #[inline]
    pub fn push(&mut self, entry: Tag) -> Result<()> {
        self.push_with_limits(entry, Limits::default())
    }

    /// Append a tag under caller-supplied resource bounds.
    pub fn push_with_limits(&mut self, entry: Tag, limits: Limits) -> Result<()> {
        self.check_add(&entry, limits)?;
        self.entries
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "designer tags",
                source,
            })?;
        self.bytes = self.bytes.checked_add(entry.bytes()?).ok_or(Error::Limit {
            resource: "designer tag aggregate bytes",
            limit: limits.aggregate_bytes,
        })?;
        self.entries.push(entry);
        Ok(())
    }

    /// Insert a tag at a source-order position under caller-supplied bounds.
    pub fn insert_with_limits(&mut self, index: usize, entry: Tag, limits: Limits) -> Result<()> {
        if index > self.entries.len() {
            return Err(Error::IndexOutOfBounds {
                index,
                len: self.entries.len(),
            });
        }
        self.check_add(&entry, limits)?;
        self.entries
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "designer tags",
                source,
            })?;
        self.bytes = self.bytes.checked_add(entry.bytes()?).ok_or(Error::Limit {
            resource: "designer tag aggregate bytes",
            limit: limits.aggregate_bytes,
        })?;
        self.entries.insert(index, entry);
        Ok(())
    }

    /// Replace one tag, returning the previous entry.
    pub fn replace_with_limits(&mut self, index: usize, entry: Tag, limits: Limits) -> Result<Tag> {
        let previous = self.entries.get(index).ok_or(Error::IndexOutOfBounds {
            index,
            len: self.entries.len(),
        })?;
        let previous_bytes = previous.bytes()?;
        let entry_bytes = entry.validate(limits)?;
        let reduced = self.bytes.checked_sub(previous_bytes).ok_or(Error::Limit {
            resource: "designer tag aggregate bytes",
            limit: limits.aggregate_bytes,
        })?;
        let bytes = reduced.checked_add(entry_bytes).ok_or(Error::Limit {
            resource: "designer tag aggregate bytes",
            limit: limits.aggregate_bytes,
        })?;
        if bytes > limits.aggregate_bytes {
            return Err(Error::Limit {
                resource: "designer tag aggregate bytes",
                limit: limits.aggregate_bytes,
            });
        }
        self.bytes = bytes;
        Ok(std::mem::replace(&mut self.entries[index], entry))
    }

    /// Remove and return one tag at a source-order position.
    pub fn remove(&mut self, index: usize) -> Result<Tag> {
        let entry = self.entries.get(index).ok_or(Error::IndexOutOfBounds {
            index,
            len: self.entries.len(),
        })?;
        self.bytes = self.bytes.checked_sub(entry.bytes()?).ok_or(Error::Limit {
            resource: "designer tag aggregate bytes",
            limit: usize::MAX,
        })?;
        Ok(self.entries.remove(index))
    }

    /// Clear this present collection without making it absent.
    #[inline]
    pub fn clear(&mut self) {
        self.entries.clear();
        self.bytes = 0;
    }

    pub(crate) fn validate(&self, limits: Limits) -> Result<()> {
        if self.entries.len() > limits.tag_count {
            return Err(Error::Limit {
                resource: "designer tag count",
                limit: limits.tag_count,
            });
        }
        let mut bytes = 0usize;
        for entry in &self.entries {
            bytes = bytes
                .checked_add(entry.validate(limits)?)
                .ok_or(Error::Limit {
                    resource: "designer tag aggregate bytes",
                    limit: limits.aggregate_bytes,
                })?;
            if bytes > limits.aggregate_bytes {
                return Err(Error::Limit {
                    resource: "designer tag aggregate bytes",
                    limit: limits.aggregate_bytes,
                });
            }
        }
        if bytes != self.bytes {
            return Err(Error::Invalid(
                "designer tag aggregate byte count is inconsistent".into(),
            ));
        }
        Ok(())
    }

    fn check_add(&self, entry: &Tag, limits: Limits) -> Result<()> {
        if self.entries.len() >= limits.tag_count {
            return Err(Error::Limit {
                resource: "designer tag count",
                limit: limits.tag_count,
            });
        }
        let entry_bytes = entry.validate(limits)?;
        let bytes = self.bytes.checked_add(entry_bytes).ok_or(Error::Limit {
            resource: "designer tag aggregate bytes",
            limit: limits.aggregate_bytes,
        })?;
        if bytes > limits.aggregate_bytes {
            return Err(Error::Limit {
                resource: "designer tag aggregate bytes",
                limit: limits.aggregate_bytes,
            });
        }
        Ok(())
    }
}

/// Package-independent Designer drawing properties.
///
/// `editable` retains whether the source explicitly supplied the attribute;
/// [`DrawingProperties::effective_editable`] applies its schema default of
/// `false`. `None` tags means no tag collection, while `Some(Tags::new())`
/// represents an explicit empty collection.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DrawingProperties {
    editable: Option<bool>,
    tags: Option<Tags>,
}

impl DrawingProperties {
    /// Construct properties with both optional members absent.
    #[inline]
    pub const fn new() -> Self {
        Self {
            editable: None,
            tags: None,
        }
    }

    /// Return the explicitly stored editable value, if present.
    #[inline]
    pub const fn editable(&self) -> Option<bool> {
        self.editable
    }

    /// Return the effective editable value, applying the schema default.
    #[inline]
    pub const fn effective_editable(&self) -> bool {
        match self.editable {
            Some(value) => value,
            None => false,
        }
    }

    /// Return the optional tag collection.
    #[inline]
    pub fn tags(&self) -> Option<&Tags> {
        self.tags.as_ref()
    }

    /// Replace the explicit editable value, retaining absence when `None`.
    #[inline]
    pub fn set_editable(&mut self, value: Option<bool>) {
        self.editable = value;
    }

    /// Replace the optional tag collection after bounded validation.
    pub fn set_tags(&mut self, value: Option<Tags>, limits: Limits) -> Result<()> {
        if let Some(tags) = &value {
            tags.validate(limits)?;
        }
        self.tags = value;
        Ok(())
    }

    /// Set the explicit editable value while building properties.
    #[must_use]
    #[inline]
    pub const fn with_editable(mut self, value: Option<bool>) -> Self {
        self.editable = value;
        self
    }

    /// Set the optional tag collection while building properties.
    pub fn with_tags(mut self, value: Option<Tags>, limits: Limits) -> Result<Self> {
        self.set_tags(value, limits)?;
        Ok(self)
    }
}

/// An extension entry that this release does not interpret.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Opaque {
    pub(super) xml: Vec<u8>,
}

impl Opaque {
    /// Borrow the exact extension bytes retained by the bounded snapshot.
    #[inline]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }

    pub(crate) fn from_wire(xml: Vec<u8>) -> Result<Self> {
        if xml.is_empty() {
            return Err(Error::Invalid("designer opaque extension is empty".into()));
        }
        if xml.len() > MAX_UNKNOWN_EXTENSION_BYTES {
            return Err(Error::Limit {
                resource: "designer opaque extension bytes",
                limit: MAX_UNKNOWN_EXTENSION_BYTES,
            });
        }
        Ok(Self { xml })
    }
}

/// A detached, lossless snapshot of one shape's optional `designElem`.
///
/// `None` from [`Snapshot::value`] preserves a present, schema-valid
/// `designElem` whose optional `val` attribute was omitted. It is distinct
/// from `Some(false)`, which is serialized as `val="false"`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(super) value: Option<bool>,
    pub(super) unknown_extensions: Vec<Opaque>,
}

impl Snapshot {
    /// Construct a snapshot with an explicit design-element value.
    #[inline]
    pub fn new(value: bool) -> Self {
        Self {
            value: Some(value),
            unknown_extensions: Vec::new(),
        }
    }

    /// Return the optional `designElem/@val` value.
    #[inline]
    pub const fn value(&self) -> Option<bool> {
        self.value
    }

    /// Borrow unrelated extension entries retained byte-for-byte.
    #[inline]
    pub fn unknown_extensions(&self) -> &[Opaque] {
        &self.unknown_extensions
    }

    /// Start a detached atomic edit of this snapshot.
    #[inline]
    pub fn edit(&self) -> crate::shape::designer::Editor {
        crate::shape::designer::Editor::new(self.clone())
    }

    pub(crate) fn from_wire(value: Option<bool>, unknown_extensions: Vec<Opaque>) -> Result<Self> {
        let snapshot = Self {
            value,
            unknown_extensions,
        };
        snapshot.validate()?;
        Ok(snapshot)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if self.unknown_extensions.len() > MAX_UNKNOWN_EXTENSIONS {
            return Err(Error::Limit {
                resource: "designer opaque extensions",
                limit: MAX_UNKNOWN_EXTENSIONS,
            });
        }
        let mut total = 0usize;
        for extension in &self.unknown_extensions {
            if extension.xml.is_empty() || extension.xml.len() > MAX_UNKNOWN_EXTENSION_BYTES {
                return Err(Error::Limit {
                    resource: "designer opaque extension bytes",
                    limit: MAX_UNKNOWN_EXTENSION_BYTES,
                });
            }
            total = total.checked_add(extension.xml.len()).ok_or(Error::Limit {
                resource: "designer opaque extension bytes",
                limit: MAX_UNKNOWN_BYTES,
            })?;
            if total > MAX_UNKNOWN_BYTES {
                return Err(Error::Limit {
                    resource: "designer opaque extension bytes",
                    limit: MAX_UNKNOWN_BYTES,
                });
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{DrawingProperties, Limits, Tag, Tags};
    use crate::Error;

    #[test]
    fn tags_preserve_empty_values_duplicates_and_order() {
        let empty = Tag::new("", "").unwrap();
        let duplicate = Tag::new("", "").unwrap();
        let tags = Tags::from_iter_with_limits([empty, duplicate], Limits::default()).unwrap();

        assert_eq!(tags.len(), 2);
        assert_eq!(tags.as_slice()[0].name(), "");
        assert_eq!(tags.as_slice()[1].value(), "");
        assert_eq!(tags.aggregate_bytes(), 0);
    }

    #[test]
    fn tag_rejects_xml_10_forbidden_characters() {
        assert!(matches!(Tag::new("bad\0", "value"), Err(Error::Invalid(_))));
        assert!(matches!(
            Tag::new("name", "bad\u{1}"),
            Err(Error::Invalid(_))
        ));
    }

    #[test]
    fn limits_apply_per_string_count_and_aggregate_bytes() {
        let tight = Limits::new()
            .with_string_bytes(2)
            .with_tag_count(1)
            .with_aggregate_bytes(3);
        assert!(matches!(
            Tag::new_with_limits("abc", "", tight),
            Err(Error::Limit {
                resource: "designer tag name",
                limit: 2,
            })
        ));

        let mut tags = Tags::new();
        tags.push_with_limits(Tag::new("a", "b").unwrap(), tight)
            .unwrap();
        assert!(matches!(
            tags.push_with_limits(Tag::new("", "").unwrap(), tight),
            Err(Error::Limit {
                resource: "designer tag count",
                limit: 1,
            })
        ));

        let aggregate = Limits::new().with_string_bytes(4).with_aggregate_bytes(3);
        assert!(matches!(
            Tags::from_iter_with_limits([Tag::new("ab", "cd").unwrap()], aggregate),
            Err(Error::Limit {
                resource: "designer tag aggregate bytes",
                limit: 3,
            })
        ));
    }

    #[test]
    fn drawing_properties_keep_absence_and_explicit_empty_distinct() {
        let mut properties = DrawingProperties::new();
        assert_eq!(properties.editable(), None);
        assert!(!properties.effective_editable());
        assert_eq!(properties.tags(), None);

        properties.set_editable(Some(true));
        properties
            .set_tags(Some(Tags::new()), Limits::default())
            .unwrap();
        assert_eq!(properties.editable(), Some(true));
        assert!(properties.effective_editable());
        assert!(properties.tags().is_some_and(Tags::is_empty));
    }
}
