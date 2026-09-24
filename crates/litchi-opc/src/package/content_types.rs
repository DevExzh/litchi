//! Source-preserving `[Content_Types].xml` structural edits.

use super::super::{OwnedElementEdit, OwnedElementUpdate, OwnedXmlPart, PackURI, ReadLimits};
use crate::Result;
use crate::content_type::validate_content_type;
use crate::error::OpcError;
use crate::limits::ReadResource;
use crate::source_backed::content_types_plan::collapse_xml_token;
use crate::source_backed::escaped_xml_attribute_len;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::sync::Arc as SharedArc;

/// One borrowed content-type override to add to a source manifest.
///
/// The part name and MIME value are validated during planning.  The final
/// manifest is not allocated until [`ContentTypesEditPlan::materialize`] is
/// called.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ContentTypeEdit<'a> {
    /// Canonical OPC part name receiving the explicit override.
    pub part_name: &'a PackURI,
    /// MIME value written to the `ContentType` attribute.
    pub content_type: &'a str,
}

/// A source-bound content-types edit admission.
///
/// Planning retains the source allocation and small lexical ranges while
/// computing exact final bytes and mapping count.  It does not build a
/// replacement XML buffer.  Callers should compare these metrics with their
/// aggregate package limits before consuming the plan with `materialize`.
#[derive(Debug)]
pub struct ContentTypesEditPlan<'source> {
    source: &'source super::OwnedContentTypes,
    inner: crate::source_backed::content_types_plan::ContentTypesPlan<'source>,
}

impl<'source> ContentTypesEditPlan<'source> {
    /// Exact final `[Content_Types].xml` byte length.
    #[must_use]
    pub const fn final_len(&self) -> usize {
        self.inner.final_len()
    }

    /// Exact final number of `Default` and `Override` mappings.
    #[must_use]
    pub const fn mapping_count(&self) -> usize {
        self.inner.mapping_count()
    }

    /// Exact final quick-xml event count, including `Event::Eof`.
    #[must_use]
    pub const fn event_count(&self) -> usize {
        self.inner.event_count()
    }

    /// Whether this plan is an exact source-preserving no-op.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.inner.is_noop()
    }

    /// Materialize one validated replacement XML buffer after caller limits
    /// have admitted the exact plan metrics.
    pub fn materialize(self, limits: ReadLimits) -> Result<super::OwnedContentTypes> {
        let source = self.source;
        limits.check(
            ReadResource::ContentTypesBytes,
            self.inner.final_len() as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            ReadResource::PartBytes,
            self.inner.final_len() as u64,
            limits.max_part_bytes(),
        )?;
        limits.check(
            ReadResource::ArchiveEntryBytes,
            self.inner.final_len() as u64,
            limits.max_archive_entry_bytes(),
        )?;
        limits.check(
            ReadResource::XmlEvents,
            self.inner.event_count() as u64,
            limits.max_xml_events() as u64,
        )?;
        limits.check(
            ReadResource::ContentTypeMappings,
            self.inner.mapping_count() as u64,
            limits.max_content_type_mappings() as u64,
        )?;
        if self.inner.is_noop() {
            return Ok(source.clone());
        }
        let bytes = self.inner.materialize_public(limits)?;
        let binding = SharedArc::new(crate::content_type::ContentTypeMap::from_xml(
            &bytes, limits,
        )?);
        let xml = OwnedXmlPart::capture_with_limits(
            source.xml.name.clone(),
            source.xml.content_type.clone(),
            SharedArc::new(bytes),
            limits,
        )?;
        Ok(super::OwnedContentTypes { xml, binding })
    }

    /// The source token used by this plan.  This is useful for exact no-op
    /// decisions without exposing its internal XML splice ranges.
    #[must_use]
    pub fn source(&self) -> &'source super::OwnedContentTypes {
        self.source
    }
}

impl super::OwnedContentTypes {
    /// Plan source-preserving additions and exact `<Override>` removals.
    ///
    /// Existing unrelated defaults, overrides, comments, processing
    /// instructions, namespace aliases, and lexical ordering remain in the
    /// source.  The returned plan computes exact final bytes without
    /// allocating candidate XML.
    pub fn plan_edit<'source>(
        &'source self,
        additions: &[ContentTypeEdit<'_>],
        removals: &[PackURI],
        limits: ReadLimits,
    ) -> Result<ContentTypesEditPlan<'source>> {
        self.plan_edit_with_defaults(additions, removals, &[], limits)
    }

    /// Plan source-preserving overrides plus selected `<Default>` mapping
    /// removals.  Default selectors use XML token-normalized, ASCII
    /// case-insensitive extension matching, while the original lexical bytes
    /// of every retained mapping remain untouched.
    pub fn plan_edit_with_defaults<'source>(
        &'source self,
        additions: &[ContentTypeEdit<'_>],
        removals: &[PackURI],
        default_removals: &[&str],
        limits: ReadLimits,
    ) -> Result<ContentTypesEditPlan<'source>> {
        limits.check(
            ReadResource::ContentTypesBytes,
            self.xml.bytes().len() as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            ReadResource::PartBytes,
            self.xml.bytes().len() as u64,
            limits.max_part_bytes(),
        )?;
        limits.check(
            ReadResource::ArchiveEntryBytes,
            self.xml.bytes().len() as u64,
            limits.max_archive_entry_bytes(),
        )?;
        preflight_content_type_edit_inputs(additions, removals, default_removals, limits)?;
        for addition in additions {
            if self.binding.override_for(addition.part_name).is_some()
                && !removals
                    .iter()
                    .any(|removal| removal.is_equivalent_to(addition.part_name))
            {
                return Err(OpcError::InvalidContentTypesManifest(
                    "content-types override already exists for the part".to_owned(),
                ));
            }
        }
        let mut owned_additions = Vec::new();
        owned_additions
            .try_reserve_exact(additions.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC content-types edit additions",
                source,
            })?;
        for addition in additions {
            owned_additions.push((
                clone_pack_uri_bounded(addition.part_name, "OPC content-types edit part names")?,
                clone_string_bounded(addition.content_type, "OPC content-types edit MIME values")?,
            ));
        }
        let mut owned_defaults = Vec::new();
        owned_defaults
            .try_reserve_exact(default_removals.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC content-types default removal selectors",
                source,
            })?;
        for value in default_removals {
            owned_defaults.push(collapse_xml_token(value)?);
        }
        let inner = crate::source_backed::content_types_plan::ContentTypesPlan::plan_public_owned(
            self.xml.bytes(),
            owned_additions,
            removals,
            owned_defaults,
            limits,
        )?;
        Ok(ContentTypesEditPlan {
            source: self,
            inner,
        })
    }
}

fn clone_string_bounded(value: &str, resource: &'static str) -> Result<String> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| OpcError::Allocation { resource, source })?;
    owned.push_str(value);
    Ok(owned)
}

fn clone_pack_uri_bounded(value: &PackURI, resource: &'static str) -> Result<PackURI> {
    let text = clone_string_bounded(value.as_str(), resource)?;
    PackURI::new(text).map_err(OpcError::InvalidPackUri)
}

/// Bound operation selectors before copying any caller-owned strings into a
/// source-backed plan.  The source scanner performs the exact escaped-output
/// admission later; this pass bounds the raw input retained by the plan and
/// applies the same per-attribute ceilings as content-type parsing.
fn preflight_content_type_edit_inputs(
    additions: &[ContentTypeEdit<'_>],
    removals: &[PackURI],
    default_removals: &[&str],
    limits: ReadLimits,
) -> Result<()> {
    if additions.len() > limits.max_content_type_mappings() {
        return Err(OpcError::ReadLimit {
            resource: ReadResource::ContentTypeMappings,
            actual: additions.len() as u64,
            maximum: limits.max_content_type_mappings() as u64,
        });
    }
    let removal_count = removals
        .len()
        .checked_add(default_removals.len())
        .ok_or_else(|| OpcError::InvalidContentTypesManifest("selector count overflows".into()))?;
    if removal_count > limits.max_content_type_mappings() {
        return Err(OpcError::ReadLimit {
            resource: ReadResource::ContentTypeMappings,
            actual: removal_count as u64,
            maximum: limits.max_content_type_mappings() as u64,
        });
    }

    let mut raw_bytes = 0usize;
    for addition in additions {
        let escaped_part_len = escaped_xml_attribute_len(addition.part_name.as_str())?;
        let escaped_content_type_len = escaped_xml_attribute_len(addition.content_type)?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            escaped_part_len as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            escaped_content_type_len as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        validate_content_type(addition.content_type).map_err(|reason| {
            OpcError::InvalidContentType {
                value: addition.content_type.to_owned(),
                reason,
            }
        })?;
        raw_bytes = raw_bytes
            .checked_add(escaped_part_len)
            .and_then(|total| total.checked_add(escaped_content_type_len))
            .ok_or_else(|| {
                OpcError::InvalidContentTypesManifest("selector bytes overflow".into())
            })?;
    }
    for part in removals {
        let length = part.as_str().len();
        limits.check(
            ReadResource::XmlAttributeBytes,
            length as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        raw_bytes = raw_bytes.checked_add(length).ok_or_else(|| {
            OpcError::InvalidContentTypesManifest("selector bytes overflow".into())
        })?;
    }
    for extension in default_removals {
        limits.check(
            ReadResource::XmlAttributeBytes,
            extension.len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        raw_bytes = raw_bytes.checked_add(extension.len()).ok_or_else(|| {
            OpcError::InvalidContentTypesManifest("selector bytes overflow".into())
        })?;
    }
    limits.check(
        ReadResource::ContentTypesBytes,
        raw_bytes as u64,
        limits.max_content_types_bytes() as u64,
    )?;
    Ok(())
}

/// Remove `<Override>` elements for the supplied part names and their
/// relationship members. The source XML is edited through validated structural
/// spans, so producer comments and lexical ordering outside those elements are
/// retained exactly.
pub(crate) fn without_part_overrides(
    source: &OwnedXmlPart,
    parts: &[PackURI],
    max_output_bytes: usize,
) -> Result<OwnedXmlPart> {
    let limits = ReadLimits::default();
    let maximum = max_output_bytes.min(limits.max_content_types_bytes());
    if source.bytes().len() > maximum {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-types source exceeds the requested output limit".into(),
        ));
    }
    if parts.len() > limits.max_content_type_mappings() {
        return Err(OpcError::InvalidContentTypesManifest(
            "too many content-types removal selectors".into(),
        ));
    }
    let mut selector_bytes = 0usize;
    for part in parts {
        // A relationship URI adds at most "/_rels" and ".rels". Charge the
        // upper bound before constructing another URI from a caller's name.
        selector_bytes = selector_bytes
            .checked_add(part.as_str().len())
            .and_then(|total| total.checked_add(part.as_str().len()))
            .and_then(|total| total.checked_add(27))
            .ok_or_else(|| {
                OpcError::InvalidContentTypesManifest(
                    "content-types removal selector size overflows".into(),
                )
            })?;
    }
    if selector_bytes > limits.max_content_types_bytes() {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-types removal selectors exceed the metadata limit".into(),
        ));
    }
    if parts.is_empty() {
        return Ok(source.clone());
    }
    let mut removals = Vec::new();
    removals
        .try_reserve(parts.len().saturating_mul(2))
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types removal names",
            source,
        })?;
    for part in parts {
        removals.push(part.clone());
        let relationships = part.rels_uri().map_err(OpcError::InvalidPackUri)?;
        removals.push(relationships);
    }
    removals
        .sort_unstable_by(|left, right| cmp_ascii_case_insensitive(left.as_str(), right.as_str()));
    removals.dedup_by(|left, right| left.is_equivalent_to(right));

    // Production callers hold an OwnedContentTypes token whose ContentTypeMap
    // already validated the namespace and attribute vocabulary.
    let mut reader = NsReader::from_reader(source.bytes());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut updates = Vec::new();
    updates
        .try_reserve_exact(removals.len())
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types removal spans",
            source,
        })?;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event()?;
        let end = reader.buffer_position() as usize;
        if end < start || end > source.bytes().len() {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-types XML event range is invalid".into(),
            ));
        }
        match event {
            Event::Start(element) => {
                if depth == 1 && element.local_name().as_ref() == b"Override" {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "non-empty content-type Overrides are unsupported for source removal"
                            .into(),
                    ));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest(
                        "content-types XML depth overflows".into(),
                    )
                })?;
            },
            Event::Empty(element) if depth == 1 && element.local_name().as_ref() == b"Override" => {
                let mut part_name = None;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| {
                        OpcError::InvalidContentTypesManifest(error.to_string())
                    })?;
                    if attribute.key.as_ref() == b"PartName" {
                        let value = attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|error| {
                                OpcError::InvalidContentTypesManifest(error.to_string())
                            })?;
                        part_name = Some(
                            PackURI::new(value.into_owned()).map_err(OpcError::InvalidPackUri)?,
                        );
                    }
                }
                if let Some(part_name) = part_name
                    && removals
                        .binary_search_by(|candidate| {
                            cmp_ascii_case_insensitive(candidate.as_str(), part_name.as_str())
                        })
                        .is_ok()
                {
                    updates.push(OwnedElementUpdate {
                        start_tag: start..end,
                        edit: OwnedElementEdit::Remove,
                    });
                }
            },
            Event::Empty(_) => {},
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest(
                        "unmatched content-types XML closing element".into(),
                    )
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(OpcError::InvalidContentTypesManifest(
            "unclosed content-types XML element".into(),
        ));
    }
    if updates.is_empty() {
        return Ok(source.clone());
    }
    source.update_elements(&updates, maximum)
}

pub(crate) fn preflight_part_overrides(
    source: &OwnedXmlPart,
    overrides: &[(&PackURI, &str)],
    max_output_bytes: usize,
) -> Result<(usize, usize)> {
    let limits = ReadLimits::default();
    let maximum = max_output_bytes.min(limits.max_content_types_bytes());
    if source.bytes().len() > maximum {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-types source exceeds the requested output limit".into(),
        ));
    }
    if overrides.len() > limits.max_content_type_mappings() {
        return Err(OpcError::InvalidContentTypesManifest(
            "too many content-type override selectors".into(),
        ));
    }
    let mut selector_bytes = 0usize;
    for (part, content_type) in overrides {
        selector_bytes = selector_bytes
            .checked_add(part.as_str().len())
            .and_then(|size| size.checked_add(content_type.len()))
            .ok_or_else(|| {
                OpcError::InvalidContentTypesManifest("content-type override size overflows".into())
            })?;
        // Check each caller-provided field against the requested input cap,
        // but do not charge a made-up per-entry XML overhead here. The exact
        // escaped output is calculated below and is the only final-byte cap.
        if part.as_str().len() > maximum || content_type.len() > maximum {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-type override selectors exceed the metadata limit".into(),
            ));
        }
        // Raw selector bytes are a strict lower bound for the final escaped
        // XML, so this aggregate admission cannot reject a candidate whose
        // exact final output fits. Unlike the former +64-per-entry estimate,
        // it avoids false failures for many small overrides while bounding
        // selector metadata before names are cloned and sorted.
        if selector_bytes > maximum {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-type override selectors exceed the metadata limit".into(),
            ));
        }
    }
    Ok((maximum, selector_bytes))
}

/// Append one or more source-preserving `<Override>` children to the OPC
/// `Types` element. The caller supplies already validated part names and
/// content types; this helper preserves the existing manifest bytes and
/// validates the resulting source part before returning it.
pub(crate) fn with_part_overrides(
    source: &OwnedXmlPart,
    overrides: &[(&PackURI, &str)],
    max_output_bytes: usize,
) -> Result<OwnedXmlPart> {
    let (maximum, _selector_bytes) = preflight_part_overrides(source, overrides, max_output_bytes)?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(overrides.len())
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types override names",
            source,
        })?;
    names.extend(overrides.iter().map(|(part, _)| (*part).clone()));
    names.sort_unstable_by(|left, right| cmp_ascii_case_insensitive(left.as_str(), right.as_str()));
    if names
        .windows(2)
        .any(|window| window[0].is_equivalent_to(&window[1]))
    {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-type override selectors contain duplicate part names".into(),
        ));
    }
    if overrides.is_empty() {
        return Ok(source.clone());
    }

    let mut reader = NsReader::from_reader(source.bytes());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_tag = None;
    let mut root_name = None;
    let mut root_close = None;
    let mut root_empty = false;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event()?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) if depth == 0 => {
                if element.local_name().as_ref() != b"Types" {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "content-types root must be Types".into(),
                    ));
                }
                root_tag = Some(start..end);
                root_name = Some(element.name().as_ref().to_vec());
                depth = 1;
            },
            Event::Empty(element) if depth == 0 => {
                if element.local_name().as_ref() != b"Types" {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "content-types root must be Types".into(),
                    ));
                }
                root_tag = Some(start..end);
                root_name = Some(element.name().as_ref().to_vec());
                root_empty = true;
                break;
            },
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest(
                        "content-types XML depth overflows".into(),
                    )
                })?;
            },
            Event::End(_) => {
                if depth == 1 {
                    root_close = Some(start..end);
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OpcError::InvalidContentTypesManifest("unbalanced content-types XML".into())
                })?;
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let root_tag = root_tag.ok_or_else(|| {
        OpcError::InvalidContentTypesManifest("content-types Types root is missing".into())
    })?;
    if !root_empty && root_close.is_none() {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-types Types root is unclosed".into(),
        ));
    }
    let root_name = root_name.ok_or_else(|| {
        OpcError::InvalidContentTypesManifest("content-types Types root name is missing".into())
    })?;
    let element_name = root_name.iter().position(|byte| *byte == b':').map_or_else(
        || b"Override".to_vec(),
        |colon| {
            let mut name = root_name[..colon].to_vec();
            name.extend_from_slice(b":Override");
            name
        },
    );
    let mut fragment_len = 0usize;
    for (part, content_type) in overrides {
        let escaped_part_len = escaped_xml_attribute_len(part.as_str())?;
        let escaped_content_type_len = escaped_xml_attribute_len(content_type)?;
        let fields = b"<"
            .len()
            .checked_add(element_name.len())
            .and_then(|size| size.checked_add(b" PartName=\"\" ContentType=\"\"/>".len()))
            .and_then(|size| size.checked_add(escaped_part_len))
            .and_then(|size| size.checked_add(escaped_content_type_len))
            .ok_or_else(|| {
                OpcError::InvalidContentTypesManifest("content-type override size overflows".into())
            })?;
        fragment_len = fragment_len.checked_add(fields).ok_or_else(|| {
            OpcError::InvalidContentTypesManifest("content-type override size overflows".into())
        })?;
    }
    let (range, suffix) = if let Some(close) = root_close.as_ref() {
        (close.start..close.start, Vec::new())
    } else {
        let mut suffix = Vec::new();
        suffix
            .try_reserve(root_name.len().saturating_add(4))
            .map_err(|source| OpcError::Allocation {
                resource: "OPC content-types root close",
                source,
            })?;
        suffix.extend_from_slice(b">");
        suffix.extend_from_slice(b"</");
        suffix.extend_from_slice(&root_name);
        suffix.push(b'>');
        (root_tag.end.saturating_sub(2)..root_tag.end, suffix)
    };
    let size = source
        .bytes()
        .len()
        .checked_sub(range.len())
        .and_then(|size| size.checked_add(fragment_len))
        .and_then(|size| size.checked_add(suffix.len()))
        .ok_or_else(|| {
            OpcError::InvalidContentTypesManifest("content-type output size overflows".into())
        })?;
    if size > maximum {
        return Err(OpcError::InvalidContentTypesManifest(
            "content-type override output exceeds the metadata limit".into(),
        ));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(size)
        .map_err(|source| OpcError::Allocation {
            resource: "OPC content-types override output",
            source,
        })?;
    bytes.extend_from_slice(&source.bytes()[..range.start]);
    if root_empty {
        bytes.extend_from_slice(&suffix[..1]);
        append_override_bytes(&mut bytes, &element_name, overrides);
        bytes.extend_from_slice(&suffix[1..]);
    } else {
        append_override_bytes(&mut bytes, &element_name, overrides);
    }
    bytes.extend_from_slice(&source.bytes()[range.end..]);
    OwnedXmlPart::capture(
        source.name.clone(),
        source.content_type.clone(),
        SharedArc::new(bytes),
    )
}

fn append_override_bytes(
    output: &mut Vec<u8>,
    element_name: &[u8],
    overrides: &[(&PackURI, &str)],
) {
    for (part, content_type) in overrides {
        output.extend_from_slice(b"<");
        output.extend_from_slice(element_name);
        output.extend_from_slice(b" PartName=\"");
        append_xml_escaped(output, part.as_str());
        output.extend_from_slice(b"\" ContentType=\"");
        append_xml_escaped(output, content_type);
        output.extend_from_slice(b"\"/>");
    }
}

fn append_xml_escaped(output: &mut Vec<u8>, value: &str) {
    for byte in value.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\'' => output.extend_from_slice(b"&apos;"),
            byte => output.push(byte),
        }
    }
}

fn cmp_ascii_case_insensitive(left: &str, right: &str) -> std::cmp::Ordering {
    for (left, right) in left.bytes().zip(right.bytes()) {
        let ordering = left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase());
        if ordering != std::cmp::Ordering::Equal {
            return ordering;
        }
    }
    left.len().cmp(&right.len())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn removes_part_and_relationship_overrides_without_losing_lexical_context() {
        let bytes = br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<!-- keep -->
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/keep.xml" ContentType="application/xml"/>
<!-- remove target -->
<Override PartName="/word/drop.xml" ContentType="application/xml"/>
<Override PartName="/word/_rels/drop.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>"#;
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(bytes.to_vec()),
        )
        .unwrap();
        let output = without_part_overrides(
            &source,
            &[PackURI::new("/word/drop.xml").unwrap()],
            1024 * 1024,
        )
        .unwrap();
        let output = std::str::from_utf8(output.bytes()).unwrap();
        assert!(output.contains("<!-- keep -->"));
        assert!(output.contains("<!-- remove target -->"));
        assert!(output.contains("/word/keep.xml"));
        assert!(!output.contains("/word/drop.xml\""));
        assert!(!output.contains("drop.xml.rels"));
    }

    #[test]
    fn matches_case_insensitive_part_names_and_refuses_nonempty_overrides() {
        let bytes = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/drop.xml" ContentType="application/xml"/></Types>"#;
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(bytes.to_vec()),
        )
        .unwrap();
        let output =
            without_part_overrides(&source, &[PackURI::new("/WORD/DROP.XML").unwrap()], 1024)
                .unwrap();
        assert!(
            !std::str::from_utf8(output.bytes())
                .unwrap()
                .contains("drop.xml")
        );

        let nonempty = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/drop.xml" ContentType="application/xml"><x/></Override></Types>"#.to_vec(),
            ),
        )
        .unwrap();
        assert!(
            without_part_overrides(&nonempty, &[PackURI::new("/word/drop.xml").unwrap()], 1024,)
                .is_err()
        );
    }

    #[test]
    fn adds_overrides_without_normalizing_unrelated_manifest_bytes() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><!-- retain --><Default Extension="xml" ContentType="application/xml"/></Types>"#.to_vec(),
            ),
        )
        .unwrap();
        let part = PackURI::new("/word/ink/ink1.xml").unwrap();
        let output =
            with_part_overrides(&source, &[(&part, "application/inkml+xml")], 4096).unwrap();
        let output = std::str::from_utf8(output.bytes()).unwrap();
        assert!(output.contains("<!-- retain -->"));
        assert!(output.contains("PartName=\"/word/ink/ink1.xml\""));
        assert!(output.contains("ContentType=\"application/inkml+xml\""));
    }

    #[test]
    fn borrowed_content_types_plan_reports_exact_mixed_output_and_removes_default() {
        let source_bytes = br#"<?xml version='1.0'?><Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
<!-- keep --> <Default Extension='xml' ContentType='application/xml'/>
<Default Extension='svg' ContentType='image/svg+xml'/>
<Override PartName='/custom/drop.bin' ContentType='application/octet-stream'/>
</Types>"#;
        let name = PackURI::new("/[Content_Types].xml").unwrap();
        let xml = OwnedXmlPart::capture(
            name,
            "application/xml".to_owned(),
            SharedArc::new(source_bytes.to_vec()),
        )
        .unwrap();
        let binding = SharedArc::new(
            crate::content_type::ContentTypeMap::from_xml(source_bytes, ReadLimits::default())
                .unwrap(),
        );
        let token = super::super::OwnedContentTypes { xml, binding };
        let added = PackURI::new("/custom/new.bin").unwrap();
        let dropped = PackURI::new("/custom/drop.bin").unwrap();
        let edits = [ContentTypeEdit {
            part_name: &added,
            content_type: "application/octet-stream",
        }];
        let plan = token
            .plan_edit_with_defaults(&edits, &[dropped], &["svg"], ReadLimits::default())
            .unwrap();
        assert_eq!(plan.mapping_count(), 2);
        let exact = plan.final_len().max(token.bytes().len());
        let limits = ReadLimits::builder()
            .max_content_types_bytes(exact)
            .unwrap()
            .max_part_bytes(exact as u64)
            .unwrap()
            .build()
            .unwrap();
        let after = token
            .plan_edit_with_defaults(
                &edits,
                &[PackURI::new("/custom/drop.bin").unwrap()],
                &["SVG"],
                limits,
            )
            .unwrap()
            .materialize(limits)
            .unwrap();
        let text = std::str::from_utf8(after.bytes()).unwrap();
        assert!(text.contains("<!-- keep -->"));
        assert!(text.contains("/custom/new.bin"));
        assert!(!text.contains("Extension='svg'"));
        assert!(!text.contains("/custom/drop.bin"));
    }

    #[test]
    fn borrowed_content_types_plan_rejects_retained_override_duplicate_but_allows_selected_replacement()
     {
        let source_bytes = br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/custom/existing.bin' ContentType='application/old'/></Types>"#;
        let name = PackURI::new("/[Content_Types].xml").unwrap();
        let xml = OwnedXmlPart::capture(
            name,
            "application/xml".to_owned(),
            SharedArc::new(source_bytes.to_vec()),
        )
        .unwrap();
        let binding = SharedArc::new(
            crate::content_type::ContentTypeMap::from_xml(source_bytes, ReadLimits::default())
                .unwrap(),
        );
        let token = super::super::OwnedContentTypes { xml, binding };
        let existing = PackURI::new("/custom/existing.bin").unwrap();
        let edit = [ContentTypeEdit {
            part_name: &existing,
            content_type: "application/new",
        }];
        assert!(token.plan_edit(&edit, &[], ReadLimits::default()).is_err());
        let after = token
            .plan_edit(
                &edit,
                std::slice::from_ref(&existing),
                ReadLimits::default(),
            )
            .unwrap()
            .materialize(ReadLimits::default())
            .unwrap();
        let text = std::str::from_utf8(after.bytes()).unwrap();
        assert!(text.contains("ContentType=\"application/new\""));
        assert!(!text.contains("application/old"));
    }

    #[test]
    fn default_removal_selector_collapses_xml_token_whitespace_but_preserves_nbsp() {
        let source_bytes = br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='svg' ContentType='image/svg+xml'/></Types>"#;
        let name = PackURI::new("/[Content_Types].xml").unwrap();
        let xml = OwnedXmlPart::capture(
            name,
            "application/xml".to_owned(),
            SharedArc::new(source_bytes.to_vec()),
        )
        .unwrap();
        let binding = SharedArc::new(
            crate::content_type::ContentTypeMap::from_xml(source_bytes, ReadLimits::default())
                .unwrap(),
        );
        let token = super::super::OwnedContentTypes { xml, binding };
        let spaced = [" \tSVG\r\n"];
        let after = token
            .plan_edit_with_defaults(&[], &[], &spaced, ReadLimits::default())
            .unwrap()
            .materialize(ReadLimits::default())
            .unwrap();
        assert!(
            !std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("Extension='svg'")
        );

        let nbsp = ["\u{00a0}svg"];
        assert!(
            token
                .plan_edit_with_defaults(&[], &[], &nbsp, ReadLimits::default())
                .is_err()
        );
        let duplicate = [" svg ", "SVG\n"];
        assert!(
            token
                .plan_edit_with_defaults(&[], &[], &duplicate, ReadLimits::default())
                .is_err()
        );
    }

    #[test]
    fn borrowed_content_types_plan_expands_prefixed_empty_root_and_rejects_one_under() {
        let source_bytes = br#"<ct:Types xmlns:ct='http://schemas.openxmlformats.org/package/2006/content-types'/>"#;
        let name = PackURI::new("/[Content_Types].xml").unwrap();
        let xml = OwnedXmlPart::capture(
            name,
            "application/xml".to_owned(),
            SharedArc::new(source_bytes.to_vec()),
        )
        .unwrap();
        let binding = SharedArc::new(
            crate::content_type::ContentTypeMap::from_xml(source_bytes, ReadLimits::default())
                .unwrap(),
        );
        let token = super::super::OwnedContentTypes { xml, binding };
        let added = PackURI::new("/new.bin").unwrap();
        let edits = [ContentTypeEdit {
            part_name: &added,
            content_type: "application/octet-stream",
        }];
        let plan = token.plan_edit(&edits, &[], ReadLimits::default()).unwrap();
        assert_eq!(plan.event_count(), 4);
        let exact = plan.final_len().max(token.bytes().len());
        let limits = ReadLimits::builder()
            .max_content_types_bytes(exact)
            .unwrap()
            .max_part_bytes(exact as u64)
            .unwrap()
            .build()
            .unwrap();
        let after = token
            .plan_edit(&edits, &[], limits)
            .unwrap()
            .materialize(limits)
            .unwrap();
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("</ct:Types>")
        );
        let under = ReadLimits::builder()
            .max_content_types_bytes(exact - 1)
            .unwrap()
            .max_part_bytes((exact - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(token.plan_edit(&edits, &[], under).is_err());
    }

    #[test]
    fn expands_self_closing_types_and_batches_overrides_as_one_root() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#
                    .to_vec(),
            ),
        )
        .unwrap();
        let first = PackURI::new("/word/ink/ink1.xml").unwrap();
        let second = PackURI::new("/word/media/ink1.png").unwrap();
        let output = with_part_overrides(
            &source,
            &[(&first, "application/inkml+xml"), (&second, "image/png")],
            4096,
        )
        .unwrap();
        let text = std::str::from_utf8(output.bytes()).unwrap();
        assert_eq!(text.matches("<Types").count(), 1);
        assert_eq!(text.matches("</Types>").count(), 1);
        assert_eq!(text.matches("<Override ").count(), 2);
        assert!(text.contains("PartName=\"/word/ink/ink1.xml\""));
        assert!(text.contains("PartName=\"/word/media/ink1.png\""));
    }

    #[test]
    fn self_closing_types_preserve_trailing_lexical_members() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/><!--tail--><?keep?>
"#
                    .to_vec(),
            ),
        )
        .unwrap();
        let part = PackURI::new("/word/ink/ink1.xml").unwrap();
        let output =
            with_part_overrides(&source, &[(&part, "application/inkml+xml")], 4096).unwrap();
        let text = std::str::from_utf8(output.bytes()).unwrap();
        assert!(text.ends_with("<!--tail--><?keep?>\n"));
    }

    #[test]
    fn preserves_a_prefixed_types_namespace_when_batching_overrides() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><!--keep--></ct:Types>"#.to_vec(),
            ),
        )
        .unwrap();
        let part = PackURI::new("/word/ink/ink1.xml").unwrap();
        let output =
            with_part_overrides(&source, &[(&part, "application/inkml+xml")], 4096).unwrap();
        let text = std::str::from_utf8(output.bytes()).unwrap();
        assert!(text.contains("<!--keep-->") && text.contains("<ct:Override "));
        assert_eq!(text.matches("<ct:Types").count(), 1);
    }

    #[test]
    fn default_empty_root_exact_cap_and_one_under_fail_before_output_allocation() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#
                    .to_vec(),
            ),
        )
        .unwrap();
        let part = PackURI::new("/word/media/vector.svg").unwrap();
        let overrides = [(&part, "image/svg+xml")];
        let unconstrained = with_part_overrides(&source, &overrides, usize::MAX).unwrap();
        let exact = unconstrained.bytes().len();
        let exact_output = with_part_overrides(&source, &overrides, exact).unwrap();
        assert_eq!(exact_output.bytes(), unconstrained.bytes());
        let under = with_part_overrides(&source, &overrides, exact - 1).unwrap_err();
        assert!(matches!(
            under,
            OpcError::InvalidContentTypesManifest(message)
                if message.contains("output exceeds")
        ));
    }

    #[test]
    fn prefixed_empty_root_exact_cap_preserves_prefix_and_typed_readback() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"/>"#
                    .to_vec(),
            ),
        )
        .unwrap();
        let part = PackURI::new("/word/media/vector.svg").unwrap();
        let overrides = [(&part, "image/svg+xml")];
        let output = with_part_overrides(&source, &overrides, usize::MAX).unwrap();
        let text = std::str::from_utf8(output.bytes()).unwrap();
        assert!(text.contains("<ct:Override "));
        assert!(text.contains("</ct:Types>"));
        let exact = output.bytes().len();
        assert!(with_part_overrides(&source, &overrides, exact).is_ok());
        assert!(with_part_overrides(&source, &overrides, exact - 1).is_err());
    }

    #[test]
    fn many_small_overrides_use_exact_escaped_cap_without_selector_overcharge() {
        let source = OwnedXmlPart::capture(
            PackURI::new("/[Content_Types].xml").unwrap(),
            "application/xml".into(),
            std::sync::Arc::new(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#.to_vec(),
            ),
        )
        .unwrap();
        let parts: Vec<_> = (0..32)
            .map(|index| PackURI::new(format!("/word/media/v{index}.svg")).unwrap())
            .collect();
        let overrides: Vec<_> = parts.iter().map(|part| (part, "image/svg+xml")).collect();
        let output = with_part_overrides(&source, &overrides, usize::MAX).unwrap();
        let exact = output.bytes().len();
        assert!(with_part_overrides(&source, &overrides, exact).is_ok());
        let under = with_part_overrides(&source, &overrides, exact - 1).unwrap_err();
        assert!(matches!(
            under,
            OpcError::InvalidContentTypesManifest(message)
                if message.contains("output exceeds")
        ));
    }
}
