//! Bounded XML codec for Slide Synchronization Data parts.

use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, process_markup_compatibility};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use super::model::{DateTime, Properties};
use crate::presentation_properties::metadata::escape_xml;
use crate::presentation_properties::metadata::is_presentationml_name;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::SourceLimits;

const MAX_PART_XML_BYTES: usize = 1024 * 1024;
const MAX_TOTAL_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_PARTS: usize = 4096;
const MAX_XML_NODES: usize = 4096;
const MAX_XML_DEPTH: usize = 32;
const MAX_SERVER_SLIDE_ID_BYTES: usize = 4096;

/// Caller-selected bounds for Slide Synchronization Data XML operations.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_total_input_bytes: usize,
    max_parts: usize,
    max_nodes: usize,
    max_depth: usize,
    max_server_slide_id_bytes: usize,
}

impl Limits {
    /// The historical Slide Synchronization safety profile.
    pub const DEFAULT: Self = Self {
        max_input_bytes: MAX_PART_XML_BYTES,
        max_output_bytes: MAX_PART_XML_BYTES,
        max_total_input_bytes: MAX_TOTAL_XML_BYTES,
        max_parts: MAX_PARTS,
        max_nodes: MAX_XML_NODES,
        max_depth: MAX_XML_DEPTH,
        max_server_slide_id_bytes: MAX_SERVER_SLIDE_ID_BYTES,
    };

    /// Construct a finite policy no wider than the format safety ceilings.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_total_input_bytes: usize,
        max_parts: usize,
        max_nodes: usize,
        max_depth: usize,
        max_server_slide_id_bytes: usize,
    ) -> Option<Self> {
        let default = Self::DEFAULT;
        if max_input_bytes == 0
            || max_input_bytes > default.max_input_bytes
            || max_output_bytes == 0
            || max_output_bytes > default.max_output_bytes
            || max_total_input_bytes == 0
            || max_total_input_bytes > default.max_total_input_bytes
            || max_parts == 0
            || max_parts > default.max_parts
            || max_nodes == 0
            || max_nodes > default.max_nodes
            || max_depth == 0
            || max_depth > default.max_depth
            || max_server_slide_id_bytes == 0
            || max_server_slide_id_bytes > default.max_server_slide_id_bytes
        {
            return None;
        }
        Some(Self {
            max_input_bytes,
            max_output_bytes,
            max_total_input_bytes,
            max_parts,
            max_nodes,
            max_depth,
            max_server_slide_id_bytes,
        })
    }

    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
    #[must_use]
    pub const fn max_total_input_bytes(self) -> usize {
        self.max_total_input_bytes
    }
    #[must_use]
    pub const fn max_parts(self) -> usize {
        self.max_parts
    }
    #[must_use]
    pub const fn max_nodes(self) -> usize {
        self.max_nodes
    }
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }
    #[must_use]
    pub const fn max_server_slide_id_bytes(self) -> usize {
        self.max_server_slide_id_bytes
    }

    /// Derive a profile from the retained physical source policy.
    pub fn from_source_limits(source: SourceLimits) -> Self {
        let default = Self::DEFAULT;
        let part = usize::try_from(source.max_part_bytes()).unwrap_or(usize::MAX);
        let total = usize::try_from(source.max_total_part_bytes()).unwrap_or(usize::MAX);
        let events = source.max_events();
        let depth = source.max_depth();
        let strings = source.max_xml_attribute_bytes();
        Self {
            max_input_bytes: default.max_input_bytes.min(part),
            max_output_bytes: default.max_output_bytes.min(part),
            max_total_input_bytes: default.max_total_input_bytes.min(total),
            max_parts: default.max_parts.min(source.max_parts()),
            max_nodes: default.max_nodes.min(events),
            max_depth: default.max_depth.min(depth),
            max_server_slide_id_bytes: default.max_server_slide_id_bytes.min(strings),
        }
    }

    /// Intersect this semantic profile with physical source bounds.
    pub(crate) fn constrained_by_source(self, source: SourceLimits) -> Self {
        let physical = Self::from_source_limits(source);
        Self {
            max_input_bytes: self.max_input_bytes.min(physical.max_input_bytes),
            max_output_bytes: self.max_output_bytes.min(physical.max_output_bytes),
            max_total_input_bytes: self
                .max_total_input_bytes
                .min(physical.max_total_input_bytes),
            max_parts: self.max_parts.min(physical.max_parts),
            max_nodes: self.max_nodes.min(physical.max_nodes),
            max_depth: self.max_depth.min(physical.max_depth),
            max_server_slide_id_bytes: self
                .max_server_slide_id_bytes
                .min(physical.max_server_slide_id_bytes),
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

impl Properties {
    /// Parse a complete Slide Synchronization Data part.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(xml: &[u8], limits: Limits) -> Result<Self> {
        if xml.len() > limits.max_input_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::InputBytes,
                xml.len(),
                limits.max_input_bytes(),
            ));
        }
        // Apply the selected MCE budget while AlternateContent is being
        // expanded.  Checking the expanded size only afterwards can already
        // have paid for an over-budget intermediate document.
        let mce_limits = mce_limits(limits);
        let xml = process_markup_compatibility(xml, &Capabilities::default(), &mce_limits)?.xml;
        if xml.len() > limits.max_input_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                xml.len(),
                limits.max_input_bytes(),
            ));
        }
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut properties = None;
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let is_empty_element = matches!(event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    nodes = increment(
                        nodes,
                        limits.max_nodes(),
                        crate::error::CodecLimitKind::XmlNodes,
                    )?;
                    depth = increment(
                        depth,
                        limits.max_depth(),
                        crate::error::CodecLimitKind::XmlDepth,
                    )?;
                    let is_root = is_presentationml_name(&namespace, element.name(), b"sldSyncPr");
                    let is_extension_list =
                        is_presentationml_name(&namespace, element.name(), b"extLst");
                    match depth {
                        1 => {
                            if properties.is_some() || closed_root {
                                return Err(invalid(
                                    "slide synchronization part has multiple root elements",
                                ));
                            }
                            if !is_root {
                                return Err(invalid(
                                    "slide synchronization part must have a sldSyncPr root",
                                ));
                            }
                            properties = Some(from_root(&element, decoder, limits)?);
                        },
                        2 if is_extension_list => {
                            let entry = properties.as_mut().ok_or_else(|| {
                                invalid("missing slide synchronization root element")
                            })?;
                            if entry.has_extension_list {
                                return Err(invalid(
                                    "duplicate slide synchronization extension list",
                                ));
                            }
                            entry.has_extension_list = true;
                        },
                        2 => {
                            return Err(invalid(
                                "unexpected child of the slide synchronization root",
                            ));
                        },
                        _ if is_root => {
                            return Err(invalid("nested slide synchronization root element"));
                        },
                        _ => {},
                    }
                    if is_empty_element {
                        depth -= 1;
                        if depth == 0 {
                            closed_root = true;
                        }
                    }
                },
                Event::Text(text) if depth > 0 => {
                    let content = text
                        .xml_content(quick_xml::XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    if depth == 1 && !content.trim().is_empty() {
                        return Err(invalid(
                            "slide synchronization root cannot carry text content",
                        ));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid slide synchronization XML nesting"))?;
                    if depth == 0 {
                        closed_root = true;
                    }
                },
                Event::Eof if depth != 0 => {
                    return Err(invalid("unterminated slide synchronization XML"));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        properties.ok_or_else(|| invalid("slide synchronization part is empty"))
    }

    /// Serialize a complete Slide Synchronization Data part.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml(&self, limits: Limits) -> Result<Vec<u8>> {
        if self.server_slide_id.len() > limits.max_server_slide_id_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::StringBytes,
                self.server_slide_id.len(),
                limits.max_server_slide_id_bytes(),
            ));
        }
        // Preflight the escaped byte count before constructing the escaped
        // string or the final output buffer.  XML entities can expand a
        // caller-provided ID substantially.
        let escaped_server_slide_id_len = escaped_xml_len(&self.server_slide_id)?;
        let server_modified_time = self.server_modified_time.to_lexical();
        let client_inserted_time = self.client_inserted_time.to_lexical();
        let prefix = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
        let root = br#"<p:sldSyncPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" serverSldId=""#;
        let server_time = br#"" serverSldModifiedTime=""#;
        let client_time = br#"" clientInsertedTime=""#;
        let suffix = if self.has_extension_list {
            br#""><p:extLst/></p:sldSyncPr>"#.as_slice()
        } else {
            br#""/>"#.as_slice()
        };
        let output_len = prefix
            .len()
            .checked_add(root.len())
            .and_then(|length| length.checked_add(escaped_server_slide_id_len))
            .and_then(|length| length.checked_add(server_time.len()))
            .and_then(|length| length.checked_add(server_modified_time.len()))
            .and_then(|length| length.checked_add(client_time.len()))
            .and_then(|length| length.checked_add(client_inserted_time.len()))
            .and_then(|length| length.checked_add(suffix.len()))
            .ok_or(Error::CapacityOverflow {
                resource: "serialized slide synchronization XML bytes",
            })?;
        if output_len > limits.max_output_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::OutputBytes,
                output_len,
                limits.max_output_bytes(),
            ));
        }
        let server_slide_id = escape_xml(&self.server_slide_id);
        if server_slide_id.len() != escaped_server_slide_id_len {
            return Err(invalid(
                "slide synchronization escaped server slide ID length changed",
            ));
        }
        let mut out = Vec::new();
        out.try_reserve_exact(output_len)
            .map_err(|source| Error::Allocation {
                resource: "slide synchronization XML",
                source,
            })?;
        out.extend_from_slice(prefix);
        out.extend_from_slice(root);
        out.extend_from_slice(server_slide_id.as_bytes());
        out.extend_from_slice(server_time);
        out.extend_from_slice(server_modified_time.as_bytes());
        out.extend_from_slice(client_time);
        out.extend_from_slice(client_inserted_time.as_bytes());
        out.extend_from_slice(suffix);
        if out.len() != output_len {
            return Err(invalid(
                "slide synchronization output length changed during emission",
            ));
        }
        let mut parse_limits = limits;
        parse_limits.max_input_bytes = parse_limits.max_input_bytes.max(out.len());
        Self::parse(&out, parse_limits)?;
        Ok(out)
    }
}

fn mce_limits(limits: Limits) -> MceLimits {
    MceLimits {
        max_input_bytes: limits.max_input_bytes(),
        max_output_bytes: limits.max_output_bytes(),
        max_depth: limits.max_depth(),
        max_namespace_bindings: 4096.min(limits.max_nodes()),
        max_directive_tokens: 4096.min(limits.max_nodes()),
        max_choices_per_alternate: 1024.min(limits.max_nodes()),
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    }
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let bytes = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length.checked_add(bytes).ok_or(Error::CapacityOverflow {
            resource: "slide synchronization escaped server slide ID",
        })
    })
}

fn from_root(element: &BytesStart<'_>, decoder: Decoder, limits: Limits) -> Result<Properties> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some() {
            continue;
        }
        match attribute.key.local_name().as_ref() {
            b"serverSldId" | b"serverSldModifiedTime" | b"clientInsertedTime" => {},
            _ => {
                return Err(invalid(format!(
                    "unexpected slide synchronization attribute '{}'",
                    String::from_utf8_lossy(attribute.key.local_name().as_ref())
                )));
            },
        }
    }
    let server_slide_id = unqualified_attribute_value(element, b"serverSldId", decoder)?
        .ok_or_else(|| invalid("missing slide synchronization serverSldId attribute"))?;
    if server_slide_id.len() > limits.max_server_slide_id_bytes() {
        return Err(limit(
            crate::error::CodecLimitKind::StringBytes,
            server_slide_id.len(),
            limits.max_server_slide_id_bytes(),
        ));
    }
    let server_modified_time = DateTime::parse(
        &unqualified_attribute_value(element, b"serverSldModifiedTime", decoder)?.ok_or_else(
            || invalid("missing slide synchronization serverSldModifiedTime attribute"),
        )?,
    )?;
    let client_inserted_time = DateTime::parse(
        &unqualified_attribute_value(element, b"clientInsertedTime", decoder)?
            .ok_or_else(|| invalid("missing slide synchronization clientInsertedTime attribute"))?,
    )?;
    Ok(Properties {
        server_slide_id,
        server_modified_time,
        client_inserted_time,
        has_extension_list: false,
    })
}

fn increment(value: usize, max: usize, kind: crate::error::CodecLimitKind) -> Result<usize> {
    let next = value
        .checked_add(1)
        .ok_or_else(|| limit(kind, usize::MAX, max))?;
    if next > max {
        return Err(limit(kind, next, max));
    }
    Ok(next)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn escaped_server_id_is_counted_before_output_admission() {
        assert_eq!(escaped_xml_len("a&<'\"").unwrap(), 1 + 5 + 4 + 6 + 6);

        let timestamp = DateTime::parse("2024-01-02T03:04:05Z").unwrap();
        let value = Properties::new("a&<'\"", timestamp.clone(), timestamp);
        let written = value.to_xml(Limits::DEFAULT).unwrap();
        assert!(
            written
                .windows(b"a&amp;&lt;&apos;&quot;".len())
                .any(|window| { window == b"a&amp;&lt;&apos;&quot;" })
        );

        let tight = Limits::new(
            Limits::DEFAULT.max_input_bytes(),
            written.len() - 1,
            Limits::DEFAULT.max_total_input_bytes(),
            Limits::DEFAULT.max_parts(),
            Limits::DEFAULT.max_nodes(),
            Limits::DEFAULT.max_depth(),
            Limits::DEFAULT.max_server_slide_id_bytes(),
        )
        .unwrap();
        assert!(value.to_xml(tight).is_err());
    }
}
