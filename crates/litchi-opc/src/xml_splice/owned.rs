//! Checked source-preserving XML publication for the eager, owned OPC package.

use super::{invalid_source, validate_source_xml};
use crate::{OpcError, PackURI, ReadLimits, Result};
use quick_xml::{events::Event, reader::NsReader};
use std::{collections::HashMap, fmt, ops::Range, sync::Arc};

const MAX_OWNED_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_EDITS: usize = 65_536;

mod elements;
mod sequence;
pub use elements::{OwnedElementEdit, OwnedElementUpdate};
pub use sequence::OwnedChildElement;

/// An unqualified attribute edit on an exact source opening tag.
/// A missing value removes the attribute; supplied values must be XML-escaped.
pub struct OwnedAttributeUpdate<'a> {
    pub start_tag: Range<usize>,
    pub name: &'a str,
    pub value: Option<&'a [u8]>,
}

impl fmt::Debug for OwnedAttributeUpdate<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("OwnedAttributeUpdate")
            .field("start_tag", &self.start_tag)
            .field("name", &self.name)
            .field("value_bytes", &self.value.map(<[u8]>::len))
            .finish()
    }
}

enum AttributeOutput<'a> {
    Value(&'a [u8]),
    Insert(&'a [u8], &'a [u8]),
    Remove,
}

struct AttributeSplice<'a> {
    range: Range<usize>,
    output: AttributeOutput<'a>,
}

struct ElementBounds {
    full: Range<usize>,
    closing_start: usize,
    empty: bool,
    name: Vec<u8>,
}

/// Validated XML captured from an owned package's trusted source or compact
/// authored part. Constructors are private; arbitrary XML cannot be marked as
/// source-preserved. Cloning retains the shared immutable source allocation.
#[derive(Clone, PartialEq, Eq)]
pub struct OwnedXmlPart {
    pub(crate) name: PackURI,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
}

impl fmt::Debug for OwnedXmlPart {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("OwnedXmlPart")
            .field("name", &self.name)
            .field("bytes", &self.bytes.len())
            .finish_non_exhaustive()
    }
}

impl OwnedXmlPart {
    pub(crate) fn check_capture_member_name(name: &PackURI, limits: ReadLimits) -> Result<()> {
        limits.check(
            crate::ReadResource::ArchiveMemberNameBytes,
            name.membername().len() as u64,
            limits.max_archive_member_name_bytes(),
        )
    }

    /// Check the member-name quota for the `.rels` URI derived from an owner
    /// without constructing that URI.  This keeps a tight caller ceiling ahead
    /// of the allocation performed by [`PackURI::rels_uri`].
    pub(crate) fn check_derived_capture_member_name(
        owner: &PackURI,
        limits: ReadLimits,
    ) -> Result<()> {
        let base_uri = owner.base_uri();
        let prefix_len = if base_uri == "/" {
            b"_rels/".len()
        } else {
            base_uri
                .len()
                .checked_sub(1)
                .and_then(|length| length.checked_add(b"/_rels/".len()))
                .ok_or_else(|| {
                    OpcError::InvalidPackUri(
                        "derived relationship member name length overflows".to_owned(),
                    )
                })?
        };
        let member_name_len = prefix_len
            .checked_add(owner.filename().len())
            .and_then(|length| length.checked_add(b".rels".len()))
            .ok_or_else(|| {
                OpcError::InvalidPackUri(
                    "derived relationship member name length overflows".to_owned(),
                )
            })?;
        limits.check(
            crate::ReadResource::ArchiveMemberNameBytes,
            member_name_len as u64,
            limits.max_archive_member_name_bytes(),
        )
    }

    /// Check the bounded resources that must be admitted before retaining or
    /// parsing one owned XML publication.  Callers that have not allocated
    /// the source/output bytes yet use this preflight to keep the fixed XML
    /// ceiling and caller part quota ahead of that work.
    pub(crate) fn check_capture_size(
        name: &PackURI,
        bytes_len: usize,
        limits: ReadLimits,
    ) -> Result<()> {
        Self::check_capture_member_name(name, limits)?;
        Self::check_capture_bytes(bytes_len, limits)
    }

    /// Check the bounded byte ceilings shared by source and canonical XML
    /// capture before any owned XML buffer is allocated.
    pub(crate) fn check_capture_bytes(bytes_len: usize, limits: ReadLimits) -> Result<()> {
        limits.check(
            crate::ReadResource::PartBytes,
            bytes_len as u64,
            limits.max_part_bytes(),
        )?;
        if bytes_len > MAX_OWNED_XML_BYTES {
            return Err(invalid_source("owned XML exceeds 32 MiB"));
        }
        Ok(())
    }

    /// Update attributes in one source scan and one output allocation.
    /// Updates must be ordered by opening-tag position; names within a tag
    /// must be unique. Namespace declarations and qualified names are refused.
    /// Output is admitted against both the caller limit and the 32 MiB ceiling.
    pub fn update_attributes(
        &self,
        updates: &[OwnedAttributeUpdate<'_>],
        max_output_bytes: usize,
    ) -> Result<Self> {
        let maximum = max_output_bytes.min(MAX_OWNED_XML_BYTES);
        if updates.len() > MAX_EDITS {
            return Err(invalid_source("too many owned XML attribute updates"));
        }
        let mut previous: Option<&Range<usize>> = None;
        let mut supplied = 0usize;
        for update in updates {
            let tag = &update.start_tag;
            if tag.start >= tag.end
                || tag.end > self.bytes.len()
                || previous.is_some_and(|previous| previous != tag && tag.start < previous.end)
            {
                return Err(invalid_source(
                    "invalid or unordered XML opening-tag updates",
                ));
            }
            previous = Some(tag);
            if update.name.len() > 4096
                || update.name == "xmlns"
                || !crate::pkgreader::is_xml_id(update.name)
            {
                return Err(invalid_source("invalid unqualified XML attribute name"));
            }
            if let Some(value) = update.value {
                supplied = supplied
                    .checked_add(value.len())
                    .ok_or_else(|| invalid_source("owned XML update size overflow"))?;
                if supplied > maximum {
                    return Err(invalid_source(
                        "owned XML attribute values exceed output limit",
                    ));
                }
                if value
                    .iter()
                    .any(|b| matches!(b, b'\'' | b'"' | b'<' | b'\t' | b'\n' | b'\r'))
                {
                    return Err(invalid_source("owned XML attribute value must be escaped"));
                }
            }
        }
        let mut splices = Vec::new();
        splices
            .try_reserve_exact(updates.len())
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML update spans",
                source,
            })?;
        let mut pending = HashMap::new();
        let mut reader = NsReader::from_reader(self.bytes.as_slice());
        let mut next = 0usize;
        while next < updates.len() {
            let start = reader.buffer_position() as usize;
            let event = reader
                .read_event()
                .map_err(|error| invalid_source(error.to_string()))?;
            let end = reader.buffer_position() as usize;
            if matches!(event, Event::Eof) || start > updates[next].start_tag.start {
                return Err(invalid_source(
                    "update does not identify a complete XML opening tag",
                ));
            }
            let empty = matches!(&event, Event::Empty(_));
            let element = match event {
                Event::Start(element) | Event::Empty(element)
                    if updates[next].start_tag == (start..end) =>
                {
                    element
                },
                _ => continue,
            };
            let first = next;
            while next < updates.len() && updates[next].start_tag == (start..end) {
                next += 1;
            }
            pending.clear();
            pending
                .try_reserve(next - first)
                .map_err(|source| OpcError::Allocation {
                    resource: "owned XML attribute names",
                    source,
                })?;
            for update in &updates[first..next] {
                if pending
                    .insert(update.name.as_bytes(), update.value)
                    .is_some()
                {
                    return Err(invalid_source("duplicate owned XML attribute update"));
                }
            }
            for attribute in element.attributes().with_checks(true) {
                let attribute = attribute.map_err(|error| invalid_source(error.to_string()))?;
                let Some(value) = pending.remove(attribute.key.as_ref()) else {
                    continue;
                };
                let value_start = (attribute.value.as_ptr() as usize)
                    .checked_sub(self.bytes.as_ptr() as usize)
                    .ok_or_else(|| invalid_source("attribute value is not source-backed"))?;
                let value_end = value_start
                    .checked_add(attribute.value.len())
                    .filter(|end| *end < self.bytes.len())
                    .ok_or_else(|| invalid_source("invalid source attribute value"))?;
                if let Some(value) = value {
                    if value != attribute.value.as_ref() {
                        splices.push(AttributeSplice {
                            range: value_start..value_end,
                            output: AttributeOutput::Value(value),
                        });
                    }
                } else {
                    let key_start = (attribute.key.as_ref().as_ptr() as usize)
                        .checked_sub(self.bytes.as_ptr() as usize)
                        .ok_or_else(|| invalid_source("attribute name is not source-backed"))?;
                    splices.push(AttributeSplice {
                        range: key_start..value_end + 1,
                        output: AttributeOutput::Remove,
                    });
                }
            }
            for update in &updates[first..next] {
                if pending.contains_key(update.name.as_bytes()) {
                    if let Some(value) = update.value {
                        let at = end - if empty { 2 } else { 1 };
                        splices.push(AttributeSplice {
                            range: at..at,
                            output: AttributeOutput::Insert(update.name.as_bytes(), value),
                        });
                    }
                }
            }
        }
        let mut size = self.bytes.len();
        for splice in &splices {
            let added = match splice.output {
                AttributeOutput::Value(value) => value.len(),
                AttributeOutput::Insert(name, value) => name
                    .len()
                    .checked_add(value.len())
                    .and_then(|size| size.checked_add(4))
                    .ok_or_else(|| invalid_source("owned XML insertion size overflow"))?,
                AttributeOutput::Remove => 0,
            };
            size = size
                .checked_sub(splice.range.len())
                .and_then(|size| size.checked_add(added))
                .ok_or_else(|| invalid_source("owned XML output size overflow"))?;
        }
        if size > maximum {
            return Err(invalid_source("owned XML output exceeds limit"));
        }
        if splices.is_empty() {
            return Ok(self.clone());
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(size)
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML attribute output",
                source,
            })?;
        let mut cursor = 0usize;
        for splice in splices {
            if splice.range.start < cursor {
                return Err(invalid_source("overlapping owned XML attribute updates"));
            }
            bytes.extend_from_slice(&self.bytes[cursor..splice.range.start]);
            match splice.output {
                AttributeOutput::Value(value) => bytes.extend_from_slice(value),
                AttributeOutput::Insert(name, value) => {
                    bytes.push(b' ');
                    bytes.extend_from_slice(name);
                    bytes.extend_from_slice(b"=\"");
                    bytes.extend_from_slice(value);
                    bytes.push(b'"');
                },
                AttributeOutput::Remove => {},
            }
            cursor = splice.range.end;
        }
        bytes.extend_from_slice(&self.bytes[cursor..]);
        validate_source_xml(&self.name, &bytes, ReadLimits::default(), None)?;
        Ok(Self {
            name: self.name.clone(),
            content_type: self.content_type.clone(),
            bytes: Arc::new(bytes),
        })
    }

    pub(crate) fn capture(
        name: PackURI,
        content_type: String,
        bytes: Arc<Vec<u8>>,
    ) -> Result<Self> {
        Self::capture_with_limits(name, content_type, bytes, ReadLimits::default())
    }

    /// Admit and validate one source XML publication without retaining an
    /// [`OwnedXmlPart`].  Source-bound callers use this preflight before an
    /// operation aggregate authorizes a later capture; it deliberately shares
    /// the exact raw-event, depth, namespace, and attribute checks performed
    /// by [`Self::capture_with_limits`].
    pub(crate) fn preflight_capture_with_limits(
        name: &PackURI,
        content_type: &str,
        bytes: &[u8],
        limits: ReadLimits,
    ) -> Result<()> {
        Self::check_capture_size(name, bytes.len(), limits)?;
        if !xml_minifier::audit::package::is_xml_part(name.as_str(), content_type) {
            return Err(invalid_source("owned XML source is not an XML part"));
        }
        validate_source_xml(name, bytes, limits, None)
    }

    pub(crate) fn capture_with_limits(
        name: PackURI,
        content_type: String,
        bytes: Arc<Vec<u8>>,
        limits: ReadLimits,
    ) -> Result<Self> {
        Self::preflight_capture_with_limits(&name, &content_type, &bytes, limits)?;
        Ok(Self {
            name,
            content_type,
            bytes,
        })
    }

    /// Borrow validated source XML without copying.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Share the validated XML allocation without copying its bytes.
    ///
    /// This is useful when retaining byte guards alongside the source token
    /// in a host patch. Mutating the returned value with `Arc::make_mut`
    /// creates a separate allocation and leaves this token unchanged. The
    /// returned bytes do not themselves carry source-replacement authority.
    #[must_use]
    pub fn shared_bytes(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.bytes)
    }

    /// Append one complete compact authored element to an exact source parent.
    /// Empty parents are expanded without changing their existing attributes.
    pub fn append_element(&self, parent_tag: Range<usize>, element: &[u8]) -> Result<Self> {
        self.edit_element(parent_tag, Some(element), true)
    }

    /// Replace an exact source element with one complete compact authored element.
    pub fn replace_element(&self, start_tag: Range<usize>, element: &[u8]) -> Result<Self> {
        self.edit_element(start_tag, Some(element), false)
    }

    /// Remove one complete source element, preserving every surrounding byte.
    pub fn remove_element(&self, start_tag: Range<usize>) -> Result<Self> {
        self.edit_element(start_tag, None, false)
    }

    fn edit_element(
        &self,
        tag: Range<usize>,
        element: Option<&[u8]>,
        append: bool,
    ) -> Result<Self> {
        let bounds = self.element_bounds(tag)?;
        if !append && element.is_some_and(|element| element == &self.bytes[bounds.full.clone()]) {
            return Ok(self.clone());
        }
        if let Some(element) = element {
            if element.len() > MAX_OWNED_XML_BYTES {
                return Err(invalid_source("authored element exceeds 32 MiB"));
            }
            let _report = xml_minifier::audit::verify_authored(
                element,
                xml_minifier::audit::Limits::default(),
            )
            .map_err(|source| OpcError::XmlPublication {
                part: self.name.to_string(),
                source,
            })?;
            validate_source_xml(&self.name, element, ReadLimits::default(), None)?;
        }
        let value = element.unwrap_or_default();
        let range = if append {
            if bounds.empty {
                bounds.closing_start..bounds.full.end
            } else {
                bounds.closing_start..bounds.closing_start
            }
        } else {
            bounds.full
        };
        let suffix_size = if append && bounds.empty {
            bounds.name.len().checked_add(4)
        } else {
            Some(0)
        }
        .ok_or_else(|| invalid_source("owned XML element name size overflow"))?;
        let size = self
            .bytes
            .len()
            .checked_sub(range.len())
            .and_then(|size| size.checked_add(value.len()))
            .and_then(|size| size.checked_add(suffix_size))
            .ok_or_else(|| invalid_source("owned XML element output size overflow"))?;
        if size > MAX_OWNED_XML_BYTES {
            return Err(invalid_source("owned XML output exceeds 32 MiB"));
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(size)
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML element output",
                source,
            })?;
        bytes.extend_from_slice(&self.bytes[..range.start]);
        if append && bounds.empty {
            bytes.push(b'>');
        }
        bytes.extend_from_slice(value);
        if append && bounds.empty {
            bytes.extend_from_slice(b"</");
            bytes.extend_from_slice(&bounds.name);
            bytes.push(b'>');
        }
        bytes.extend_from_slice(&self.bytes[range.end..]);
        validate_source_xml(&self.name, &bytes, ReadLimits::default(), None)?;
        Ok(Self {
            name: self.name.clone(),
            content_type: self.content_type.clone(),
            bytes: Arc::new(bytes),
        })
    }

    fn element_bounds(&self, tag: Range<usize>) -> Result<ElementBounds> {
        if tag.start >= tag.end || tag.end > self.bytes.len() {
            return Err(invalid_source("invalid XML opening-tag range"));
        }
        let mut reader = NsReader::from_reader(self.bytes.as_slice());
        let mut depth = 0usize;
        let mut owner = None;
        loop {
            let start = usize::try_from(reader.buffer_position())
                .map_err(|error| invalid_source(error.to_string()))?;
            let event = reader
                .read_event()
                .map_err(|error| invalid_source(error.to_string()))?;
            let end = usize::try_from(reader.buffer_position())
                .map_err(|error| invalid_source(error.to_string()))?;
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if tag == (start..end) {
                        let qname = element.name();
                        let mut name = Vec::new();
                        name.try_reserve_exact(qname.as_ref().len())
                            .map_err(|source| OpcError::Allocation {
                                resource: "owned XML element name",
                                source,
                            })?;
                        name.extend_from_slice(qname.as_ref());
                        if empty {
                            return Ok(ElementBounds {
                                full: tag,
                                closing_start: end - 2,
                                empty: true,
                                name,
                            });
                        }
                        owner = Some((depth, name));
                    }
                    if !empty {
                        depth += 1;
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid_source("unbalanced XML element source"))?;
                    if owner
                        .as_ref()
                        .is_some_and(|(owner_depth, _)| *owner_depth == depth)
                    {
                        let (_, name) =
                            owner.ok_or_else(|| invalid_source("missing XML element owner"))?;
                        return Ok(ElementBounds {
                            full: tag.start..end,
                            closing_start: start,
                            empty: false,
                            name,
                        });
                    }
                },
                Event::Eof => {
                    return Err(invalid_source(
                        "range does not identify a complete XML opening tag",
                    ));
                },
                _ => {},
            }
        }
    }

    /// Add an absent unqualified attribute to one exact source opening tag.
    /// The range must cover the complete Start or Empty token. Namespace
    /// declarations and duplicate attributes are refused. The value is already
    /// XML-escaped, with the same rules as [`Self::replace_attributes`].
    pub fn insert_unqualified_attribute(
        &self,
        start_tag: Range<usize>,
        name: &str,
        escaped_value: &[u8],
    ) -> Result<Self> {
        if name.len() > 4096 || name == "xmlns" || !crate::pkgreader::is_xml_id(name) {
            return Err(invalid_source("invalid unqualified XML attribute name"));
        }
        if start_tag.start >= start_tag.end || start_tag.end > self.bytes.len() {
            return Err(invalid_source("invalid XML opening-tag range"));
        }
        let size = self
            .bytes
            .len()
            .checked_add(name.len())
            .and_then(|size| size.checked_add(escaped_value.len()))
            .and_then(|size| size.checked_add(4))
            .ok_or_else(|| invalid_source("owned XML insertion size overflow"))?;
        if size > MAX_OWNED_XML_BYTES {
            return Err(invalid_source("owned XML output exceeds 32 MiB"));
        }
        if escaped_value
            .iter()
            .any(|b| matches!(b, b'\'' | b'"' | b'<' | b'\t' | b'\n' | b'\r'))
        {
            return Err(invalid_source("owned XML attribute value must be escaped"));
        }
        let mut reader = NsReader::from_reader(self.bytes.as_slice());
        let insertion = loop {
            let start = usize::try_from(reader.buffer_position())
                .map_err(|error| invalid_source(error.to_string()))?;
            let event = reader
                .read_event()
                .map_err(|error| invalid_source(error.to_string()))?;
            let end = usize::try_from(reader.buffer_position())
                .map_err(|error| invalid_source(error.to_string()))?;
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) if start_tag == (start..end) => {
                    for attribute in element.attributes().with_checks(true) {
                        let attribute =
                            attribute.map_err(|error| invalid_source(error.to_string()))?;
                        if attribute.key.as_ref() == name.as_bytes() {
                            return Err(invalid_source("XML attribute already exists"));
                        }
                    }
                    break end - if empty { 2 } else { 1 };
                },
                Event::Eof => {
                    return Err(invalid_source(
                        "range does not identify a complete XML opening tag",
                    ));
                },
                _ => {},
            }
        };
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(size)
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML attribute insertion",
                source,
            })?;
        bytes.extend_from_slice(&self.bytes[..insertion]);
        bytes.push(b' ');
        bytes.extend_from_slice(name.as_bytes());
        bytes.extend_from_slice(b"=\"");
        bytes.extend_from_slice(escaped_value);
        bytes.push(b'"');
        bytes.extend_from_slice(&self.bytes[insertion..]);
        validate_source_xml(&self.name, &bytes, ReadLimits::default(), None)?;
        Ok(Self {
            name: self.name.clone(),
            content_type: self.content_type.clone(),
            bytes: Arc::new(bytes),
        })
    }

    /// Replace complete attribute value spans, retaining all other bytes.
    /// Ranges must be ordered, nonoverlapping, and cover precisely one source
    /// attribute value each. Replacement values must already be XML-escaped;
    /// literal quotes, angle brackets opening markup, and XML whitespace
    /// controls are rejected. Empty replacement values are supported.
    /// This validates XML safety, not format-specific attribute semantics.
    pub fn replace_attributes(&self, edits: &[(Range<usize>, Vec<u8>)]) -> Result<Self> {
        if edits.is_empty() {
            return Ok(self.clone());
        }
        if edits.len() > MAX_EDITS {
            return Err(invalid_source("too many owned XML attribute edits"));
        }
        let mut removed = 0usize;
        let mut added = 0usize;
        let mut cursor = 0usize;
        for (range, value) in edits {
            if range.start < cursor || range.end < range.start || range.end > self.bytes.len() {
                return Err(invalid_source("invalid owned XML attribute replacement"));
            }
            removed = removed
                .checked_add(range.len())
                .ok_or_else(|| invalid_source("owned XML removed size overflow"))?;
            added = added
                .checked_add(value.len())
                .ok_or_else(|| invalid_source("owned XML replacement size overflow"))?;
            if added > MAX_OWNED_XML_BYTES {
                return Err(invalid_source("owned XML replacement bytes exceed 32 MiB"));
            }
            if value
                .iter()
                .any(|b| matches!(b, b'\'' | b'"' | b'<' | b'\t' | b'\n' | b'\r'))
            {
                return Err(invalid_source("owned XML attribute value must be escaped"));
            }
            cursor = range.end;
        }
        let size = self
            .bytes
            .len()
            .checked_sub(removed)
            .and_then(|value| value.checked_add(added))
            .ok_or_else(|| invalid_source("owned XML output size overflow"))?;
        if size > MAX_OWNED_XML_BYTES {
            return Err(invalid_source("owned XML output exceeds 32 MiB"));
        }
        let mut reader = NsReader::from_reader(self.bytes.as_slice());
        let mut checked = 0;
        loop {
            match reader
                .read_event()
                .map_err(|error| invalid_source(error.to_string()))?
            {
                Event::Start(element) | Event::Empty(element) => {
                    for attribute in element.attributes().with_checks(true) {
                        let attribute =
                            attribute.map_err(|error| invalid_source(error.to_string()))?;
                        let raw = attribute.value.as_ref();
                        let start = (raw.as_ptr() as usize)
                            .checked_sub(self.bytes.as_ptr() as usize)
                            .ok_or_else(|| invalid_source("attribute is not source-backed"))?;
                        let end = start
                            .checked_add(raw.len())
                            .ok_or_else(|| invalid_source("attribute range overflow"))?;
                        if checked < edits.len() && edits[checked].0 == (start..end) {
                            checked += 1;
                        }
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if checked != edits.len() {
            return Err(invalid_source(
                "replacement does not cover a complete source attribute value",
            ));
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(size)
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML attribute output",
                source,
            })?;
        cursor = 0;
        for (range, value) in edits {
            bytes.extend_from_slice(&self.bytes[cursor..range.start]);
            bytes.extend_from_slice(value);
            cursor = range.end;
        }
        bytes.extend_from_slice(&self.bytes[cursor..]);
        validate_source_xml(&self.name, &bytes, ReadLimits::default(), None)?;
        Ok(Self {
            name: self.name.clone(),
            content_type: self.content_type.clone(),
            bytes: Arc::new(bytes),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{BlobPart, OpcPackage, PackageWriter};

    fn source(xml: &[u8]) -> OpcPackage {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer.write_stored("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/part.xml" ContentType="application/xml"/></Types>"#).unwrap();
        writer.write_stored("part.xml", xml).unwrap();
        OpcPackage::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap()
    }

    #[test]
    fn shared_xml_bytes_retain_allocation_without_exposing_token_mutation() {
        let xml = b"<r><original/></r>";
        let package = source(xml);
        let token = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        let retained = token.shared_bytes();
        let mut editable = token.shared_bytes();
        assert!(Arc::ptr_eq(&retained, &editable));
        assert_eq!(retained.as_ptr(), token.bytes().as_ptr());

        Arc::make_mut(&mut editable).clear();
        assert_eq!(token.bytes(), xml);
        assert_eq!(retained.as_slice(), xml);
        assert!(editable.is_empty());

        drop(token);
        drop(package);
        assert_eq!(retained.as_slice(), xml);
    }

    #[test]
    fn child_sequences_preserve_parent_slots_and_retained_subtree_context() {
        let xml = "<r xmlns:p='urn:p'><fixed/><!--a--><p:a v = '1'><p:inner/></p:a><!--b--><p:b/><!--end--></r>";
        let package = source(xml.as_bytes());
        let before = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        let parent = 0.."<r xmlns:p='urn:p'>".len();
        let a = xml.find("<p:a").unwrap();
        let b = xml.find("<p:b").unwrap();
        let tags = [a..a + "<p:a v = '1'>".len(), b..b + "<p:b/>".len()];
        let children = [
            OwnedChildElement::Retained(1),
            OwnedChildElement::Authored(b"<new/>"),
            OwnedChildElement::Retained(0),
            OwnedChildElement::Retained(1),
        ];
        let expected = "<r xmlns:p='urn:p'><fixed/><!--a--><p:b/><!--b--><new/><p:a v = '1'><p:inner/></p:a><p:b/><!--end--></r>";
        assert_eq!(
            before
                .replace_child_sequence(parent.clone(), &tags, &children, expected.len())
                .unwrap()
                .bytes(),
            expected.as_bytes()
        );
        assert!(
            before
                .replace_child_sequence(parent.clone(), &tags, &children, expected.len() - 1)
                .is_err()
        );
        let noop = before
            .replace_child_sequence(
                parent.clone(),
                &tags,
                &[
                    OwnedChildElement::Retained(0),
                    OwnedChildElement::Retained(1),
                ],
                xml.len(),
            )
            .unwrap();
        assert!(Arc::ptr_eq(&noop.bytes, &before.bytes));
        let removed = before
            .replace_child_sequence(parent, &tags, &[], xml.len())
            .unwrap();
        assert_eq!(
            removed.bytes(),
            b"<r xmlns:p='urn:p'><fixed/><!--a--><!--b--><!--end--></r>"
        );
    }

    #[test]
    #[allow(
        clippy::single_range_in_vec_init,
        reason = "These are XML source span lists, not sequences of integers."
    )]
    fn child_sequences_reject_cross_parent_and_unowned_content() {
        let xml = b"<r><parent><child><inner/></child></parent><other/></r>";
        let package = source(xml);
        let before = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        let parent = 3..11;
        for tags in [
            vec![18..26],
            vec![43..51],
            vec![11..17],
            vec![12..18],
            vec![11..18, 18..26],
        ] {
            assert!(
                before
                    .replace_child_sequence(parent.clone(), &tags, &[], 1024)
                    .is_err(),
                "{tags:?}"
            );
        }
        for fragment in [
            &b"<p:unbound/>"[..],
            b"<a/> <b/>",
            b"<a x = '1'/>",
            b"<!DOCTYPE a><a/>",
        ] {
            assert!(
                before
                    .replace_child_sequence(
                        parent.clone(),
                        &[11..18],
                        &[OwnedChildElement::Authored(fragment)],
                        1024
                    )
                    .is_err()
            );
        }
        assert!(
            before
                .replace_child_sequence(parent, &[11..18], &[OwnedChildElement::Retained(1)], 1024)
                .is_err()
        );
        assert_eq!(before.bytes(), xml);
    }

    #[test]
    fn child_sequence_append_expands_empty_parent_and_enforces_exact_size() {
        let xml = b"<r><p:empty xmlns:p='urn:p' /></r>";
        let package = source(xml);
        let before = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        let parent = 3..xml.len() - 4;
        let children = [
            OwnedChildElement::Authored(b"<a/>"),
            OwnedChildElement::Authored(b"<b/>"),
        ];
        let expected = b"<r><p:empty xmlns:p='urn:p' ><a/><b/></p:empty></r>";
        assert_eq!(
            before
                .replace_child_sequence(parent.clone(), &[], &children, expected.len())
                .unwrap()
                .bytes(),
            expected
        );
        assert!(
            before
                .replace_child_sequence(parent.clone(), &[], &children, expected.len() - 1)
                .is_err()
        );
        assert!(Arc::ptr_eq(
            &before.bytes,
            &before
                .replace_child_sequence(parent, &[], &[], xml.len())
                .unwrap()
                .bytes
        ));
    }

    #[test]
    fn batched_element_updates_preserve_source_and_expand_empty_parents_once() {
        let xml =
            "<root a = 'keep'><!--before--><old><inner/></old><p:empty xmlns:p='urn:p' /></root>";
        let package = source(xml.as_bytes());
        let before = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        let old = xml.find("<old>").unwrap();
        let empty = xml.find("<p:empty").unwrap();
        let empty_end = xml[empty..].find("/>").unwrap() + empty + 2;
        let updates = [
            OwnedElementUpdate {
                start_tag: 0.."<root a = 'keep'>".len(),
                edit: OwnedElementEdit::AppendChild(b"<last/>"),
            },
            OwnedElementUpdate {
                start_tag: old..old + 5,
                edit: OwnedElementEdit::InsertBefore(b"<first/>"),
            },
            OwnedElementUpdate {
                start_tag: old..old + 5,
                edit: OwnedElementEdit::Remove,
            },
            OwnedElementUpdate {
                start_tag: empty..empty_end,
                edit: OwnedElementEdit::AppendChild(b"<a/>"),
            },
            OwnedElementUpdate {
                start_tag: empty..empty_end,
                edit: OwnedElementEdit::AppendChild(b"<b/>"),
            },
        ];
        let expected = "<root a = 'keep'><!--before--><first/><p:empty xmlns:p='urn:p' ><a/><b/></p:empty><last/></root>";
        let changed = before.update_elements(&updates, expected.len()).unwrap();
        assert_eq!(changed.bytes(), expected.as_bytes());
        assert!(
            before
                .update_elements(&updates, expected.len() - 1)
                .is_err()
        );
        let noop = before
            .update_elements(
                &[OwnedElementUpdate {
                    start_tag: 0.."<root a = 'keep'>".len(),
                    edit: OwnedElementEdit::Replace(xml.as_bytes()),
                }],
                xml.len(),
            )
            .unwrap();
        assert!(Arc::ptr_eq(&before.bytes, &noop.bytes));
        assert_eq!(before.bytes(), xml.as_bytes());
    }

    #[test]
    fn batched_element_updates_reject_overlaps_bad_targets_and_unowned_fragments() {
        let xml = b"<root><child/></root>";
        let package = source(xml);
        let before = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        for updates in [
            vec![
                OwnedElementUpdate {
                    start_tag: 0..6,
                    edit: OwnedElementEdit::Remove,
                },
                OwnedElementUpdate {
                    start_tag: 6..14,
                    edit: OwnedElementEdit::Replace(b"<new/>"),
                },
            ],
            vec![
                OwnedElementUpdate {
                    start_tag: 6..14,
                    edit: OwnedElementEdit::Remove,
                },
                OwnedElementUpdate {
                    start_tag: 0..6,
                    edit: OwnedElementEdit::AppendChild(b"<new/>"),
                },
            ],
            vec![OwnedElementUpdate {
                start_tag: 0..5,
                edit: OwnedElementEdit::Remove,
            }],
            vec![OwnedElementUpdate {
                start_tag: 1..6,
                edit: OwnedElementEdit::Remove,
            }],
            vec![
                OwnedElementUpdate {
                    start_tag: 6..14,
                    edit: OwnedElementEdit::Remove,
                },
                OwnedElementUpdate {
                    start_tag: 6..14,
                    edit: OwnedElementEdit::AppendChild(b"<new/>"),
                },
            ],
        ] {
            assert!(
                before.update_elements(&updates, 1024).is_err(),
                "{updates:?}"
            );
        }
        for fragment in [
            &b"<a/> <b/>"[..],
            b"<p:unbound/>",
            b"<a x = 'noncompact'/>",
            b"<!DOCTYPE a><a/>",
            b"<a>&unknown;</a>",
            b"<a>",
        ] {
            assert!(
                before
                    .update_elements(
                        &[OwnedElementUpdate {
                            start_tag: 6..14,
                            edit: OwnedElementEdit::Replace(fragment),
                        }],
                        1024
                    )
                    .is_err(),
                "{fragment:?}"
            );
        }
        assert_eq!(before.bytes(), xml);
    }

    #[test]
    fn batched_attribute_updates_preserve_source_and_admit_exact_output_size() {
        let tag = "<root a = 'old' remove = \"x\" keep='> />'>";
        let xml = format!("{tag}<!--keep--><child a=\"\"/></root>");
        let package = source(xml.as_bytes());
        let before = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        let child = xml.find("<child").unwrap();
        let updates = [
            OwnedAttributeUpdate {
                start_tag: 0..tag.len(),
                name: "remove",
                value: None,
            },
            OwnedAttributeUpdate {
                start_tag: 0..tag.len(),
                name: "a",
                value: Some(b"&amp;new"),
            },
            OwnedAttributeUpdate {
                start_tag: 0..tag.len(),
                name: "added",
                value: Some(b"yes"),
            },
            OwnedAttributeUpdate {
                start_tag: child..child + "<child a=\"\"/>".len(),
                name: "a",
                value: Some(b"child"),
            },
        ];
        let expected = "<root a = '&amp;new'  keep='> />' added=\"yes\"><!--keep--><child a=\"child\"/></root>";
        let changed = before.update_attributes(&updates, expected.len()).unwrap();
        assert_eq!(changed.bytes(), expected.as_bytes());
        assert!(
            before
                .update_attributes(&updates, expected.len() - 1)
                .is_err()
        );
        let noop = before
            .update_attributes(
                &[
                    OwnedAttributeUpdate {
                        start_tag: 0..tag.len(),
                        name: "a",
                        value: Some(b"old"),
                    },
                    OwnedAttributeUpdate {
                        start_tag: 0..tag.len(),
                        name: "absent",
                        value: None,
                    },
                ],
                xml.len(),
            )
            .unwrap();
        assert!(Arc::ptr_eq(&noop.bytes, &before.bytes));
    }

    #[test]
    fn batched_attribute_updates_reject_ambiguous_spans_names_and_values() {
        let tag = "<root a='old'>";
        let xml = format!("{tag}<child/></root>");
        let package = source(xml.as_bytes());
        let before = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        for name in ["xmlns", "xmlns:q", "q:a", "", "bad name"] {
            assert!(
                before
                    .update_attributes(
                        &[OwnedAttributeUpdate {
                            start_tag: 0..tag.len(),
                            name,
                            value: Some(b"x")
                        },],
                        1024
                    )
                    .is_err()
            );
        }
        for value in [b"'".as_slice(), b"<child/>", b"&unknown;", b"&#0;"] {
            assert!(
                before
                    .update_attributes(
                        &[OwnedAttributeUpdate {
                            start_tag: 0..tag.len(),
                            name: "a",
                            value: Some(value)
                        },],
                        1024
                    )
                    .is_err()
            );
        }
        assert!(
            before
                .update_attributes(
                    &[
                        OwnedAttributeUpdate {
                            start_tag: 0..tag.len(),
                            name: "a",
                            value: None
                        },
                        OwnedAttributeUpdate {
                            start_tag: 0..tag.len(),
                            name: "a",
                            value: Some(b"x")
                        },
                    ],
                    1024
                )
                .is_err()
        );
        for range in [1..tag.len(), 0..tag.len() - 1, xml.len() - 7..xml.len()] {
            assert!(
                before
                    .update_attributes(
                        &[OwnedAttributeUpdate {
                            start_tag: range,
                            name: "a",
                            value: None
                        },],
                        1024
                    )
                    .is_err()
            );
        }
        let child = tag.len()..tag.len() + "<child/>".len();
        assert!(
            before
                .update_attributes(
                    &[
                        OwnedAttributeUpdate {
                            start_tag: child,
                            name: "a",
                            value: Some(b"x")
                        },
                        OwnedAttributeUpdate {
                            start_tag: 0..tag.len(),
                            name: "a",
                            value: Some(b"x")
                        },
                    ],
                    1024
                )
                .is_err()
        );
    }

    #[test]
    fn owned_xml_attribute_proof_preserves_ingress_and_supports_inverse_publication() {
        let xml = b"<root value = 'old'><!--kept--><child/></root>";
        let mut package = source(xml);
        let name = PackURI::new("/part.xml").unwrap();
        let before = package.source_xml_part(&name).unwrap();
        let at = memchr::memmem::find(xml, b"old").unwrap();
        let after = before
            .replace_attributes(&[(at..at + 3, b"new&amp;value".to_vec())])
            .unwrap();
        package
            .try_replace_owned_xml_part(xml, after.clone())
            .unwrap();
        let bytes = PackageWriter::to_bytes(&package).unwrap();
        let reopened = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            reopened.get_part(&name).unwrap().blob(),
            b"<root value = 'new&amp;value'><!--kept--><child/></root>"
        );
        assert!(
            package
                .try_replace_owned_xml_part(xml, before.clone())
                .is_err()
        );
        package
            .try_replace_owned_xml_part(after.bytes(), before.clone())
            .unwrap();
        assert_eq!(
            OpcPackage::from_bytes(&PackageWriter::to_bytes(&package).unwrap())
                .unwrap()
                .get_part(&name)
                .unwrap()
                .blob(),
            xml
        );
        package.remove_part(&name);
        package.try_add_owned_xml_part(before).unwrap();
        assert_eq!(
            OpcPackage::from_bytes(&PackageWriter::to_bytes(&package).unwrap())
                .unwrap()
                .get_part(&name)
                .unwrap()
                .blob(),
            xml
        );
    }

    #[test]
    fn owned_xml_proof_rejects_arbitrary_authored_bytes_and_attribute_injection() {
        let name = PackURI::new("/part.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            name.clone(),
            "application/xml".into(),
            b"<root value = 'old'/>".to_vec(),
        )));
        assert!(package.source_xml_part(&name).is_err());
        package
            .get_part_mut(&name)
            .unwrap()
            .set_blob(b"<root value=\"old\"/>".to_vec());
        let before = package.source_xml_part(&name).unwrap();
        let at = memchr::memmem::find(before.bytes(), b"old").unwrap();
        for value in [
            b"\" injected=\"yes".as_slice(),
            b"<markup/>",
            b"&unknown;",
            b"&#0;",
            b"\n",
        ] {
            assert!(
                before
                    .replace_attributes(&[(at..at + 3, value.to_vec())])
                    .is_err()
            );
        }
        for range in [at..at + 1, 0..4, usize::MAX..usize::MAX, at + 3..at] {
            assert!(
                before
                    .replace_attributes(&[(range, b"x".to_vec())])
                    .is_err()
            );
        }
        let empty = before
            .replace_attributes(&[(at..at + 3, Vec::new())])
            .unwrap();
        assert_eq!(
            empty
                .replace_attributes(&[(at..at, b"old".to_vec())])
                .unwrap()
                .bytes(),
            before.bytes()
        );
        assert!(
            before
                .replace_attributes(&[(at..at + 3, vec![b'x'; MAX_OWNED_XML_BYTES])])
                .is_err()
        );
    }
    #[test]
    fn owned_xml_publication_obeys_signature_policy() {
        let xml = b"<root value='old'/>";
        let mut package = source(xml);
        let name = PackURI::new("/part.xml").unwrap();
        let before = package.source_xml_part(&name).unwrap();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
                .into(),
            "_xmlsignatures/origin.sigs".into(),
            "rIdSignature".into(),
            false,
        );
        let at = memchr::memmem::find(xml, b"old").unwrap();
        let after = before
            .replace_attributes(&[(at..at + 3, b"new".to_vec())])
            .unwrap();
        assert!(matches!(
            package.try_replace_owned_xml_part(xml, after),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert!(matches!(
            package.try_add_owned_xml_part(before),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert_eq!(package.get_part(&name).unwrap().blob(), xml);
    }
    #[test]
    fn owned_xml_inserts_attributes_without_rewriting_the_parent_tag() {
        for tag in [
            "<s:root xmlns:s='urn:root' hint = '/> >' />",
            "<s:root xmlns:s='urn:root' hint = '/> >'>",
        ] {
            let xml = if tag.ends_with("/>") {
                tag.to_owned()
            } else {
                format!("{tag}<!--keep--></s:root>")
            };
            let package = source(xml.as_bytes());
            let before = package
                .source_xml_part(&PackURI::new("/part.xml").unwrap())
                .unwrap();
            let after = before
                .insert_unqualified_attribute(0..tag.len(), "version", b"7")
                .unwrap();
            let position = tag.len() - if tag.ends_with("/>") { 2 } else { 1 };
            assert_eq!(
                after.bytes(),
                format!("{} version=\"7\"{}", &xml[..position], &xml[position..]).as_bytes()
            );
            for name in ["hint", "xmlns", "s:version", "", "a b"] {
                assert!(
                    before
                        .insert_unqualified_attribute(0..tag.len(), name, b"7")
                        .is_err()
                );
            }
            assert!(
                before
                    .insert_unqualified_attribute(1..tag.len(), "version", b"7")
                    .is_err()
            );
            assert!(
                before
                    .insert_unqualified_attribute(0..tag.len(), "version", b"&undeclared;")
                    .is_err()
            );
            assert!(
                before
                    .insert_unqualified_attribute(0..tag.len(), "version", b"\" bad=\"yes")
                    .is_err()
            );
        }
    }

    #[test]
    fn owned_xml_element_edits_preserve_source_context_and_audit_new_markup() {
        let tag = "<s:root xmlns:s='urn:root' marker = 'keep'>";
        let child = "<s:child a = '1'>";
        let xml =
            format!("{tag}{child}text<!--inside--></s:child><!--between--><s:empty /></s:root>");
        let package = source(xml.as_bytes());
        let before = package
            .source_xml_part(&PackURI::new("/part.xml").unwrap())
            .unwrap();
        let appended = before.append_element(0..tag.len(), b"<new/>").unwrap();
        assert_eq!(
            appended.bytes(),
            xml.replace("</s:root>", "<new/></s:root>").as_bytes()
        );
        let start = tag.len();
        let child_tag = start..start + child.len();
        let replaced = before
            .replace_element(child_tag.clone(), b"<new/>")
            .unwrap();
        let whole_child = format!("{child}text<!--inside--></s:child>");
        assert_eq!(
            replaced.bytes(),
            xml.replace(&whole_child, "<new/>").as_bytes()
        );
        let same = before
            .replace_element(child_tag.clone(), whole_child.as_bytes())
            .unwrap();
        assert!(Arc::ptr_eq(&same.bytes, &before.bytes));
        let removed = before.remove_element(child_tag.clone()).unwrap();
        assert_eq!(removed.bytes(), xml.replace(&whole_child, "").as_bytes());
        assert!(
            before
                .replace_element(child_tag, b"<new value = '1'/>")
                .is_err()
        );
        assert!(before.remove_element(0..tag.len()).is_err());
        assert!(
            before
                .append_element(0..tag.len(), b"<new/><second/>")
                .is_err()
        );
        let empty_at = xml.find("<s:empty />").unwrap();
        let expanded = before
            .append_element(empty_at..empty_at + "<s:empty />".len(), b"<new/>")
            .unwrap();
        assert_eq!(
            expanded.bytes(),
            xml.replace("<s:empty />", "<s:empty ><new/></s:empty>")
                .as_bytes()
        );
    }

    #[test]
    fn bounded_owned_xml_capture_checks_part_quota_before_malformed_source() {
        let name = PackURI::new("/part.xml").unwrap();
        let malformed = Arc::new(b"<broken".to_vec());
        let under = ReadLimits::builder()
            .max_part_bytes((malformed.len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            OwnedXmlPart::capture_with_limits(
                name.clone(),
                "application/xml".to_owned(),
                Arc::clone(&malformed),
                under,
            ),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::PartBytes,
                actual,
                maximum,
            }) if actual == malformed.len() as u64 && maximum == (malformed.len() - 1) as u64
        ));

        let exact = ReadLimits::builder()
            .max_part_bytes(malformed.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            OwnedXmlPart::capture_with_limits(name, "application/xml".to_owned(), malformed, exact,),
            Err(OpcError::XmlError(_))
        ));
    }

    #[test]
    fn bounded_owned_xml_capture_checks_member_name_and_fixed_size_before_xml() {
        let name = PackURI::new("/part.xml").unwrap();
        let member_name_bytes = name.membername().len();
        let exact = ReadLimits::builder()
            .max_archive_member_name_bytes(member_name_bytes as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            OwnedXmlPart::capture_with_limits(
                name.clone(),
                "application/xml".to_owned(),
                Arc::new(b"<part/>".to_vec()),
                exact,
            )
            .is_ok()
        );

        let under = ReadLimits::builder()
            .max_archive_member_name_bytes((member_name_bytes - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            OwnedXmlPart::capture_with_limits(
                name.clone(),
                "application/xml".to_owned(),
                Arc::new(b"<broken".to_vec()),
                under,
            ),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::ArchiveMemberNameBytes,
                actual,
                maximum,
            }) if actual == member_name_bytes as u64 && maximum == (member_name_bytes - 1) as u64
        ));

        assert!(matches!(
            OwnedXmlPart::check_capture_size(
                &name,
                MAX_OWNED_XML_BYTES + 1,
                ReadLimits::default(),
            ),
            Err(OpcError::SourceBackedOverlayUnavailable { reason })
                if reason.contains("32 MiB")
        ));
    }
}
