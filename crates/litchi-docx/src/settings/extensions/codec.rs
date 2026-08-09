#![expect(
    clippy::expect_used,
    reason = "the invariant is established immediately before extraction"
)]
#![expect(
    clippy::match_same_arms,
    reason = "separate arms document distinct OOXML grammar cases"
)]
#![expect(
    clippy::needless_pass_by_value,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Namespace-aware XML codec for direct Word settings extensions.

use super::super::support::{invalid, xml_error};
use super::super::{
    MAX_SETTINGS_XML_BYTES, MAX_SETTINGS_XML_DEPTH, MAX_SETTINGS_XML_NODES, STRICT_WORD_NAMESPACE,
    TRANSITIONAL_WORD_NAMESPACE,
};
use super::model::{
    DocumentId, Extension, Extensions, Guid, OnOff, OpaqueExtension, WORD_2010_NAMESPACE,
    WORD_2012_NAMESPACE,
};
use super::validation::validate_opaque_xml;
use crate::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::str;

const TRANSITIONAL_WORD_NAMESPACE_TEXT: &str =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    ChartTrackingRefBased,
    ParagraphContextId,
    SourceDocumentId,
    ConflictMode,
    DiscardImageEditingData,
    DefaultImageDpi,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ByteRange {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct KnownRange {
    kind: Kind,
    range: ByteRange,
}

#[derive(Debug, Default)]
struct Layout {
    root_empty: Option<ByteRange>,
    root_close_start: Option<usize>,
    known: Vec<KnownRange>,
}

impl OpaqueExtension {
    /// Validate and retain one complete direct-child XML element.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        validate_opaque_xml(&xml)?;
        Ok(Self { xml })
    }

    fn trusted(xml: Vec<u8>) -> Self {
        Self { xml }
    }
}

impl Extensions {
    /// Parse direct settings extensions from a complete `w:settings` part.
    pub(crate) fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_SETTINGS_XML_BYTES {
            return Err(invalid(format!(
                "settings XML exceeds {MAX_SETTINGS_XML_BYTES} bytes"
            )));
        }

        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        reader.config_mut().trim_text(false);
        let mut value = Self::new();
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut saw_root = false;
        let mut pending: Option<Extension> = None;

        loop {
            let event_start = position(&reader)?;
            let event = reader
                .read_event()
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned();
            let event_end = position(&reader)?;
            let resolver = reader.resolver().clone();

            match event {
                Event::Start(element) => {
                    count_node(&mut nodes)?;
                    if depth == 0 {
                        validate_root(
                            &resolver.resolve_element(element.name()).0,
                            &element,
                            saw_root,
                        )?;
                        saw_root = true;
                        depth = 1;
                        continue;
                    }

                    if depth == 1 {
                        let (namespace, _) = resolver.resolve_element(element.name());
                        if let Some(kind) =
                            extension_kind(&namespace, element.local_name().as_ref())
                        {
                            let extension =
                                parse_known(kind, &element, reader.decoder(), &resolver)?;
                            pending = Some(extension);
                            depth = 2;
                        } else if is_wordprocessing_namespace(&namespace) {
                            depth = 2;
                        } else {
                            let bindings = active_bindings(&reader);
                            let raw = capture_unknown(
                                &mut reader,
                                xml,
                                event_start,
                                event_end,
                                depth,
                                &mut nodes,
                            )?;
                            let raw = make_self_contained(raw, &element, &bindings)?;
                            value.push(Extension::Unknown(OpaqueExtension::trusted(raw)))?;
                        }
                        continue;
                    }

                    if pending.is_some() {
                        return Err(invalid(
                            "typed Word settings extensions cannot contain child elements",
                        ));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("settings XML nesting is too deep"))?;
                    if depth > MAX_SETTINGS_XML_DEPTH {
                        return Err(invalid(format!(
                            "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                        )));
                    }
                },
                Event::Empty(element) => {
                    count_node(&mut nodes)?;
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("settings XML nesting is too deep"))?;
                    if child_depth > MAX_SETTINGS_XML_DEPTH {
                        return Err(invalid(format!(
                            "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                        )));
                    }
                    if depth == 0 {
                        validate_root(
                            &resolver.resolve_element(element.name()).0,
                            &element,
                            saw_root,
                        )?;
                        saw_root = true;
                    } else if depth == 1 {
                        let (namespace, _) = resolver.resolve_element(element.name());
                        if let Some(kind) =
                            extension_kind(&namespace, element.local_name().as_ref())
                        {
                            value.push(parse_known(
                                kind,
                                &element,
                                reader.decoder(),
                                &resolver,
                            )?)?;
                        } else if !is_wordprocessing_namespace(&namespace) {
                            let raw = copy_range(xml, event_start, event_end)?;
                            let raw =
                                make_self_contained(raw, &element, &active_bindings(&reader))?;
                            value.push(Extension::Unknown(OpaqueExtension::trusted(raw)))?;
                        }
                    } else if pending.is_some() {
                        return Err(invalid(
                            "typed Word settings extensions cannot contain child elements",
                        ));
                    }
                },
                Event::End(_) => {
                    if let Some(extension) = pending.take() {
                        if depth != 2 {
                            return Err(invalid("invalid typed settings extension nesting"));
                        }
                        value.push(extension)?;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid settings XML nesting"))?;
                },
                Event::Text(text) => {
                    if pending.is_some() {
                        return Err(invalid(
                            "typed Word settings extensions cannot contain text",
                        ));
                    }
                    if depth == 0
                        && !text
                            .decode()
                            .map_err(|error| xml_error(error.to_string()))?
                            .trim()
                            .is_empty()
                    {
                        return Err(invalid("settings XML has text outside its root"));
                    }
                },
                Event::CData(_) | Event::GeneralRef(_) if pending.is_some() => {
                    return Err(invalid(
                        "typed Word settings extensions cannot contain character data",
                    ));
                },
                Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                    return Err(invalid("settings XML has content outside its root"));
                },
                Event::Comment(_) => {},
                Event::Decl(_) if !saw_root => {},
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid(
                        "settings XML cannot contain a DTD or processing instruction",
                    ));
                },
                Event::Eof => break,
                Event::CData(_) | Event::Decl(_) | Event::GeneralRef(_) => {},
            }
        }

        if !saw_root {
            return Err(invalid("settings part has no settings root"));
        }
        if depth != 0 || pending.is_some() {
            return Err(invalid("unterminated Word settings XML"));
        }
        value.validate()?;
        Ok(value)
    }

    /// Serialize direct extension children in their current order.
    ///
    /// # Panics
    ///
    /// Panics if an internal writer invariant is violated.
    pub fn to_xml(&self, word_prefix: &str) -> String {
        let mut output = String::new();
        let mut wrote_word_2010_namespace = false;
        let mut wrote_word_2012_namespace = false;
        let mut wrote_word_namespace = false;

        for extension in &self.values {
            match extension {
                Extension::ChartTrackingRefBased(value) => write_on_off(
                    &mut output,
                    "w15",
                    WORD_2012_NAMESPACE,
                    "chartTrackingRefBased",
                    word_prefix,
                    *value,
                    &mut wrote_word_2012_namespace,
                    &mut wrote_word_namespace,
                ),
                Extension::DocumentId(DocumentId::ParagraphContext(value)) => {
                    let value = format_hex(*value);
                    write_w14_element(
                        &mut output,
                        "docId",
                        Some(("w14:val", value.as_str())),
                        &mut wrote_word_2010_namespace,
                    );
                },
                Extension::DocumentId(DocumentId::Source(value)) => {
                    let value = value.as_ref().map(ToString::to_string);
                    write_w15_element(
                        &mut output,
                        "docId",
                        value.as_deref(),
                        &mut wrote_word_2012_namespace,
                    );
                },
                Extension::ConflictMode(value) => write_w14_on_off(
                    &mut output,
                    "conflictMode",
                    *value,
                    &mut wrote_word_2010_namespace,
                ),
                Extension::DiscardImageEditingData(value) => write_w14_on_off(
                    &mut output,
                    "discardImageEditingData",
                    *value,
                    &mut wrote_word_2010_namespace,
                ),
                Extension::DefaultImageDpi(value) => {
                    let value = value.to_string();
                    write_w14_element(
                        &mut output,
                        "defaultImageDpi",
                        Some(("w14:val", value.as_str())),
                        &mut wrote_word_2010_namespace,
                    );
                },
                Extension::Unknown(value) => {
                    output.push_str(str::from_utf8(&value.xml).expect("validated opaque XML"));
                },
            }
        }
        output
    }
}

/// Rewrite only the modeled direct settings extensions.
///
/// The source ranges of every unchanged child, including unknown extension
/// XML and ordinary Word settings, are copied verbatim.  Changed typed
/// children use deterministic canonical XML; additions are appended at the
/// settings-root boundary so the source ordering of existing children remains
/// stable.
pub(crate) fn rewrite(xml: &[u8], next: &Extensions) -> Result<Vec<u8>> {
    if xml.len() > MAX_SETTINGS_XML_BYTES {
        return Err(invalid(format!(
            "settings XML exceeds {MAX_SETTINGS_XML_BYTES} bytes"
        )));
    }
    next.validate()?;
    let current = Extensions::parse(xml)?;
    if current == *next {
        return Ok(xml.to_vec());
    }

    let layout = locate_layout(xml)?;
    let mut replacements = Vec::new();
    for entry in &layout.known {
        let before = find_kind(&current, entry.kind).ok_or_else(|| {
            invalid("settings extension layout does not match its semantic projection")
        })?;
        let after = find_kind(next, entry.kind);
        let replacement = match after {
            None => None,
            Some(after) if after == before => Some(None),
            Some(after) => Some(Some(render_extension(after)?)),
        };
        replacements.push((entry.range, replacement));
    }

    let mut additions = Vec::new();
    for extension in &next.values {
        let Some(kind) = extension_kind_value(extension) else {
            continue;
        };
        if !layout.known.iter().any(|entry| entry.kind == kind) {
            additions.push(render_extension(extension)?);
        }
    }

    let additions_len = additions.iter().try_fold(0usize, |total, value| {
        total
            .checked_add(value.len())
            .ok_or_else(|| invalid("settings extension rewrite size overflows usize"))
    })?;
    let mut output_len = xml.len();
    for (range, replacement) in &replacements {
        match replacement {
            None => {
                output_len = output_len
                    .checked_sub(
                        range.end.checked_sub(range.start).ok_or_else(|| {
                            invalid("settings extension range has an invalid length")
                        })?,
                    )
                    .ok_or_else(|| invalid("settings extension rewrite size underflows usize"))?;
            },
            Some(Some(replacement)) => {
                output_len = output_len
                    .checked_sub(
                        range.end.checked_sub(range.start).ok_or_else(|| {
                            invalid("settings extension range has an invalid length")
                        })?,
                    )
                    .and_then(|value| value.checked_add(replacement.len()))
                    .ok_or_else(|| invalid("settings extension rewrite size overflows usize"))?;
            },
            Some(None) => {},
        }
    }
    output_len = output_len
        .checked_add(additions_len)
        .ok_or_else(|| invalid("settings extension rewrite size overflows usize"))?;
    if !additions.is_empty() && layout.root_close_start.is_none() {
        let root_empty = layout
            .root_empty
            .ok_or_else(|| invalid("settings root has no extension insertion point"))?;
        let root = xml
            .get(root_empty.start..root_empty.end)
            .ok_or_else(|| invalid("settings root empty range is outside the source"))?;
        let name_end = root
            .iter()
            .position(|byte| matches!(*byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))
            .ok_or_else(|| invalid("settings root empty element has no qualified name"))?;
        let name_len = name_end
            .checked_sub(1)
            .ok_or_else(|| invalid("settings root empty element has no qualified name"))?;
        output_len = output_len
            .checked_add(
                name_len
                    .checked_add(3)
                    .ok_or_else(|| invalid("settings extension rewrite size overflows usize"))?,
            )
            .ok_or_else(|| invalid("settings extension rewrite size overflows usize"))?;
    }
    if output_len > MAX_SETTINGS_XML_BYTES {
        return Err(invalid(format!(
            "settings XML exceeds {MAX_SETTINGS_XML_BYTES} bytes after extension edit"
        )));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "Word settings extension rewrite",
            source,
        })?;

    let mut cursor = 0usize;
    for (range, replacement) in replacements {
        output.extend_from_slice(
            xml.get(cursor..range.start)
                .ok_or_else(|| invalid("settings extension range starts outside the source"))?,
        );
        match replacement {
            None => {},
            Some(None) => output.extend_from_slice(
                xml.get(range.start..range.end)
                    .ok_or_else(|| invalid("settings extension range ends outside the source"))?,
            ),
            Some(Some(replacement)) => output.extend_from_slice(&replacement),
        }
        cursor = range.end;
    }

    if additions.is_empty() {
        output
            .extend_from_slice(xml.get(cursor..).ok_or_else(|| {
                invalid("settings extension source cursor is outside the source")
            })?);
        return Ok(output);
    }

    if let Some(close_start) = layout.root_close_start {
        output.extend_from_slice(
            xml.get(cursor..close_start)
                .ok_or_else(|| invalid("settings root close range is outside the source"))?,
        );
        for addition in additions {
            output.extend_from_slice(&addition);
        }
        output.extend_from_slice(
            xml.get(close_start..)
                .ok_or_else(|| invalid("settings root close range is outside the source"))?,
        );
        return Ok(output);
    }

    let root_empty = layout
        .root_empty
        .ok_or_else(|| invalid("settings root has no extension insertion point"))?;
    if cursor != root_empty.start {
        return Err(invalid(
            "settings root empty range does not match source cursor",
        ));
    }
    let root = xml
        .get(root_empty.start..root_empty.end)
        .ok_or_else(|| invalid("settings root empty range is outside the source"))?;
    let slash = root
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| invalid("settings root empty element has no closing slash"))?;
    let name_end = root
        .iter()
        .position(|byte| matches!(*byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))
        .ok_or_else(|| invalid("settings root empty element has no qualified name"))?;
    let name = root
        .get(1..name_end)
        .ok_or_else(|| invalid("settings root empty element name is outside the source"))?;
    output.extend_from_slice(&root[..slash]);
    output.push(b'>');
    for addition in additions {
        output.extend_from_slice(&addition);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.push(b'>');
    output.extend_from_slice(
        xml.get(root_empty.end..)
            .ok_or_else(|| invalid("settings root suffix is outside the source"))?,
    );
    Ok(output)
}

fn locate_layout(xml: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader.config_mut().trim_text(false);

    let mut layout = Layout::default();
    let mut stack = Vec::<Option<(Kind, usize)>>::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;

    loop {
        let event_start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();

        match event {
            Event::Start(element) => {
                if depth == 0 {
                    validate_root(
                        &resolver.resolve_element(element.name()).0,
                        &element,
                        root_seen,
                    )?;
                    root_seen = true;
                }
                let kind = if depth == 1 {
                    let (namespace, _) = resolver.resolve_element(element.name());
                    extension_kind(&namespace, element.local_name().as_ref())
                } else {
                    None
                };
                stack.push(kind.map(|kind| (kind, event_start)));
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("settings XML nesting is too deep"))?;
                if depth > MAX_SETTINGS_XML_DEPTH {
                    return Err(invalid(format!(
                        "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                    )));
                }
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("settings XML nesting is too deep"))?;
                if child_depth > MAX_SETTINGS_XML_DEPTH {
                    return Err(invalid(format!(
                        "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                    )));
                }
                if depth == 0 {
                    validate_root(
                        &resolver.resolve_element(element.name()).0,
                        &element,
                        root_seen,
                    )?;
                    root_seen = true;
                    root_closed = true;
                    layout.root_empty = Some(ByteRange {
                        start: event_start,
                        end: event_end,
                    });
                } else if depth == 1 {
                    let (namespace, _) = resolver.resolve_element(element.name());
                    if let Some(kind) = extension_kind(&namespace, element.local_name().as_ref()) {
                        layout.known.push(KnownRange {
                            kind,
                            range: ByteRange {
                                start: event_start,
                                end: event_end,
                            },
                        });
                    }
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("settings XML has an unexpected end element"));
                }
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("settings XML has an unexpected end element"))?;
                if let Some((kind, start)) = frame {
                    layout.known.push(KnownRange {
                        kind,
                        range: ByteRange {
                            start,
                            end: event_end,
                        },
                    });
                }
                if depth == 1 {
                    layout.root_close_start = Some(event_start);
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::Text(text) if depth == 0 => {
                if !text
                    .decode()
                    .map_err(|error| xml_error(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid("settings XML has text outside its root"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(invalid("settings XML has content outside its root"));
            },
            Event::Decl(_) if !root_seen => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "settings XML cannot contain a DTD or processing instruction",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !root_seen || !root_closed || depth != 0 || !stack.is_empty() {
        return Err(invalid("settings XML is not one complete root element"));
    }
    layout.known.sort_by_key(|entry| entry.range.start);
    Ok(layout)
}

fn extension_kind_value(extension: &Extension) -> Option<Kind> {
    match extension {
        Extension::ChartTrackingRefBased(_) => Some(Kind::ChartTrackingRefBased),
        Extension::DocumentId(DocumentId::ParagraphContext(_)) => Some(Kind::ParagraphContextId),
        Extension::DocumentId(DocumentId::Source(_)) => Some(Kind::SourceDocumentId),
        Extension::ConflictMode(_) => Some(Kind::ConflictMode),
        Extension::DiscardImageEditingData(_) => Some(Kind::DiscardImageEditingData),
        Extension::DefaultImageDpi(_) => Some(Kind::DefaultImageDpi),
        Extension::Unknown(_) => None,
    }
}

fn find_kind(extensions: &Extensions, kind: Kind) -> Option<&Extension> {
    extensions
        .values
        .iter()
        .find(|extension| extension_kind_value(extension) == Some(kind))
}

fn render_extension(extension: &Extension) -> Result<Vec<u8>> {
    let mut output = String::new();
    let mut wrote_word_2010_namespace = false;
    let mut wrote_word_2012_namespace = false;
    let mut wrote_word_namespace = false;
    match extension {
        Extension::ChartTrackingRefBased(value) => write_on_off(
            &mut output,
            "w15",
            WORD_2012_NAMESPACE,
            "chartTrackingRefBased",
            "w",
            *value,
            &mut wrote_word_2012_namespace,
            &mut wrote_word_namespace,
        ),
        Extension::DocumentId(DocumentId::ParagraphContext(value)) => {
            let value = format_hex(*value);
            write_w14_element(
                &mut output,
                "docId",
                Some(("w14:val", value.as_str())),
                &mut wrote_word_2010_namespace,
            );
        },
        Extension::DocumentId(DocumentId::Source(value)) => {
            let value = value.as_ref().map(ToString::to_string);
            write_w15_element(
                &mut output,
                "docId",
                value.as_deref(),
                &mut wrote_word_2012_namespace,
            );
        },
        Extension::ConflictMode(value) => write_w14_on_off(
            &mut output,
            "conflictMode",
            *value,
            &mut wrote_word_2010_namespace,
        ),
        Extension::DiscardImageEditingData(value) => write_w14_on_off(
            &mut output,
            "discardImageEditingData",
            *value,
            &mut wrote_word_2010_namespace,
        ),
        Extension::DefaultImageDpi(value) => {
            let value = value.to_string();
            write_w14_element(
                &mut output,
                "defaultImageDpi",
                Some(("w14:val", value.as_str())),
                &mut wrote_word_2010_namespace,
            );
        },
        Extension::Unknown(value) => output.push_str(
            str::from_utf8(&value.xml)
                .map_err(|_source_error| invalid("opaque XML is not UTF-8"))?,
        ),
    }
    Ok(output.into_bytes())
}

fn extension_kind(namespace: &ResolveResult<'_>, local_name: &[u8]) -> Option<Kind> {
    let is_word_2010 = bound_namespace(namespace, WORD_2010_NAMESPACE.as_bytes());
    let is_word_2012 = bound_namespace(namespace, WORD_2012_NAMESPACE.as_bytes());
    match (is_word_2010, is_word_2012, local_name) {
        (false, true, b"chartTrackingRefBased") => Some(Kind::ChartTrackingRefBased),
        (true, false, b"docId") => Some(Kind::ParagraphContextId),
        (false, true, b"docId") => Some(Kind::SourceDocumentId),
        (true, false, b"conflictMode") => Some(Kind::ConflictMode),
        (true, false, b"discardImageEditingData") => Some(Kind::DiscardImageEditingData),
        (true, false, b"defaultImageDpi") => Some(Kind::DefaultImageDpi),
        _ => None,
    }
}

fn parse_known(
    kind: Kind,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Extension> {
    match kind {
        Kind::ChartTrackingRefBased => Ok(Extension::ChartTrackingRefBased(parse_on_off(
            element,
            decoder,
            resolver,
            &[TRANSITIONAL_WORD_NAMESPACE, STRICT_WORD_NAMESPACE],
        )?)),
        Kind::ParagraphContextId => {
            let value = required_attribute(
                element,
                decoder,
                resolver,
                &[WORD_2010_NAMESPACE.as_bytes()],
                "w14:docId val",
            )?;
            let value = parse_hex(value)?;
            Ok(Extension::DocumentId(DocumentId::paragraph_context(value)?))
        },
        Kind::SourceDocumentId => {
            let value = optional_attribute(
                element,
                decoder,
                resolver,
                &[WORD_2012_NAMESPACE.as_bytes()],
                "w15:docId val",
            )?
            .map(|value| Guid::parse(value.as_str()))
            .transpose()?;
            Ok(Extension::DocumentId(DocumentId::source(value)))
        },
        Kind::ConflictMode => Ok(Extension::ConflictMode(parse_on_off(
            element,
            decoder,
            resolver,
            &[WORD_2010_NAMESPACE.as_bytes()],
        )?)),
        Kind::DiscardImageEditingData => Ok(Extension::DiscardImageEditingData(parse_on_off(
            element,
            decoder,
            resolver,
            &[WORD_2010_NAMESPACE.as_bytes()],
        )?)),
        Kind::DefaultImageDpi => Ok(Extension::DefaultImageDpi(
            required_attribute(
                element,
                decoder,
                resolver,
                &[WORD_2010_NAMESPACE.as_bytes()],
                "w14:defaultImageDpi val",
            )?
            .trim()
            .parse::<i32>()
            .map_err(|_source_error| invalid("invalid w14:defaultImageDpi val"))?,
        )),
    }
}

fn parse_on_off(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    namespaces: &[&[u8]],
) -> Result<OnOff> {
    optional_attribute(
        element,
        decoder,
        resolver,
        namespaces,
        "settings on/off val",
    )?
    .map(|value| match value.as_str() {
        "true" | "on" | "1" => Ok(true),
        "false" | "off" | "0" => Ok(false),
        _ => Err(invalid(format!(
            "invalid Word settings on/off value '{value}'"
        ))),
    })
    .transpose()
    .map(OnOff::new)
}

fn required_attribute(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    namespaces: &[&[u8]],
    description: &str,
) -> Result<String> {
    optional_attribute(element, decoder, resolver, namespaces, description)?
        .ok_or_else(|| invalid(format!("{description} attribute is required")))
}

fn optional_attribute(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    namespaces: &[&[u8]],
    description: &str,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if attribute.key.local_name().as_ref() != b"val" {
            return Err(invalid(format!("unexpected attribute on {description}")));
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let accepted = namespaces
            .iter()
            .any(|expected| bound_namespace(&namespace, expected));
        if !accepted {
            return Err(invalid(format!(
                "{description} has an attribute in the wrong namespace"
            )));
        }
        if value.is_some() {
            return Err(invalid(format!("duplicate {description} attribute")));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn parse_hex(value: String) -> Result<u32> {
    let value = value.trim();
    if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid(
            "Word settings docId must contain exactly eight hexadecimal digits",
        ));
    }
    u32::from_str_radix(value, 16).map_err(|_source_error| {
        invalid("Word settings docId contains an invalid hexadecimal value")
    })
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<()> {
    if saw_root
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"settings"
    {
        return Err(invalid(
            "settings part has an invalid or trailing root element",
        ));
    }
    Ok(())
}

fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    bound_namespace(namespace, TRANSITIONAL_WORD_NAMESPACE)
        || bound_namespace(namespace, STRICT_WORD_NAMESPACE)
}

fn bound_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == expected
    )
}

fn count_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| invalid("settings XML node counter overflow"))?;
    if *nodes > MAX_SETTINGS_XML_NODES {
        return Err(invalid(format!(
            "settings XML exceeds {MAX_SETTINGS_XML_NODES} nodes"
        )));
    }
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source_error| invalid("settings XML offset does not fit usize"))
}

fn active_bindings(reader: &NsReader<&[u8]>) -> Vec<(Option<Vec<u8>>, Vec<u8>)> {
    reader
        .resolver()
        .bindings()
        .map(|(prefix, namespace)| {
            let prefix = match prefix {
                PrefixDeclaration::Default => None,
                PrefixDeclaration::Named(prefix) => Some(prefix.to_vec()),
            };
            (prefix, namespace.as_ref().to_vec())
        })
        .collect()
}

fn make_self_contained(
    xml: Vec<u8>,
    element: &BytesStart<'_>,
    bindings: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Vec<u8>> {
    let mut declared = HashSet::<Option<Vec<u8>>>::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        if let Some(prefix) = attribute.key.as_namespace_binding() {
            declared.insert(match prefix {
                PrefixDeclaration::Default => None,
                PrefixDeclaration::Named(prefix) => Some(prefix.to_vec()),
            });
        }
    }

    let additions = bindings
        .iter()
        .filter(|(prefix, namespace)| !namespace.is_empty() && !declared.contains(prefix))
        .collect::<Vec<_>>();
    if additions.is_empty() {
        return Ok(xml);
    }

    let end = start_tag_end(&xml)?;
    let insert = if xml.get(
        end.checked_sub(1)
            .ok_or_else(|| invalid("opaque settings extension has an empty root start tag"))?,
    ) == Some(&b'/')
    {
        end - 1
    } else {
        end
    };
    let added_bytes = additions.iter().try_fold(0usize, |total, entry| {
        let prefix_len = entry.0.as_ref().map_or(0, Vec::len);
        total
            .checked_add(10 + prefix_len + entry.1.len())
            .ok_or_else(|| invalid("opaque settings namespace declaration size overflow"))
    })?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len().saturating_add(added_bytes))
        .map_err(|source| Error::Allocation {
            resource: "opaque settings extension",
            source,
        })?;
    output.extend_from_slice(&xml[..insert]);
    for entry in additions {
        output.extend_from_slice(b" xmlns");
        if let Some(prefix) = entry.0.as_ref() {
            output.push(b':');
            output.extend_from_slice(prefix);
        }
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(&entry.1);
        output.push(b'\"');
    }
    output.extend_from_slice(&xml[insert..]);
    Ok(output)
}

fn start_tag_end(xml: &[u8]) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in xml.iter().copied().enumerate() {
        match (quote, byte) {
            (None, b'\'' | b'\"') => quote = Some(byte),
            (Some(current), byte) if current == byte => quote = None,
            (None, b'>') => return Ok(index),
            _ => {},
        }
    }
    Err(invalid(
        "opaque settings extension has an unterminated root start tag",
    ))
}

fn capture_unknown(
    reader: &mut NsReader<&[u8]>,
    xml: &[u8],
    start: usize,
    _end: usize,
    parent_depth: usize,
    nodes: &mut usize,
) -> Result<Vec<u8>> {
    let mut depth = 1usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?
            .into_owned();
        let end = position(reader)?;
        match event {
            Event::Start(_) => {
                count_node(nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("settings XML nesting is too deep"))?;
                if parent_depth
                    .checked_add(depth)
                    .is_none_or(|value| value > MAX_SETTINGS_XML_DEPTH)
                {
                    return Err(invalid(format!(
                        "settings XML exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                    )));
                }
            },
            Event::Empty(_) => count_node(nodes)?,
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid unknown settings extension nesting"))?;
                if depth == 0 {
                    return copy_range(xml, start, end);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "unknown settings extensions cannot contain a DTD or processing instruction",
                ));
            },
            Event::Eof => return Err(invalid("unterminated unknown settings extension")),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn copy_range(xml: &[u8], start: usize, end: usize) -> Result<Vec<u8>> {
    let value = xml
        .get(start..end)
        .ok_or_else(|| invalid("settings XML element range is invalid"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation {
            resource: "opaque settings extension",
            source,
        })?;
    output.extend_from_slice(value);
    Ok(output)
}

fn format_hex(value: u32) -> String {
    format!("{value:08X}")
}

fn write_on_off(
    output: &mut String,
    prefix: &str,
    namespace: &str,
    local_name: &str,
    word_prefix: &str,
    value: OnOff,
    wrote_namespace: &mut bool,
    wrote_word_namespace: &mut bool,
) {
    let word_prefix = if word_prefix.is_empty() {
        "w"
    } else {
        word_prefix
    };
    output.push('<');
    output.push_str(prefix);
    output.push(':');
    output.push_str(local_name);
    write_namespace(output, prefix, namespace, wrote_namespace);
    if let Some(value) = value.authored() {
        write_namespace(
            output,
            word_prefix,
            TRANSITIONAL_WORD_NAMESPACE_TEXT,
            wrote_word_namespace,
        );
        output.push(' ');
        output.push_str(word_prefix);
        output.push_str(":val=\"");
        output.push_str(if value { "true" } else { "false" });
        output.push('"');
    }
    output.push_str("/>");
}

fn write_w14_on_off(output: &mut String, local_name: &str, value: OnOff, wrote: &mut bool) {
    output.push_str("<w14:");
    output.push_str(local_name);
    write_namespace(output, "w14", WORD_2010_NAMESPACE, wrote);
    if let Some(value) = value.authored() {
        output.push_str(" w14:val=\"");
        output.push_str(if value { "true" } else { "false" });
        output.push('"');
    }
    output.push_str("/>");
}

fn write_w14_element(
    output: &mut String,
    local_name: &str,
    attribute: Option<(&str, &str)>,
    wrote: &mut bool,
) {
    output.push_str("<w14:");
    output.push_str(local_name);
    write_namespace(output, "w14", WORD_2010_NAMESPACE, wrote);
    if let Some((name, value)) = attribute {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(value);
        output.push('"');
    }
    output.push_str("/>");
}

fn write_w15_element(output: &mut String, local_name: &str, value: Option<&str>, wrote: &mut bool) {
    output.push_str("<w15:");
    output.push_str(local_name);
    write_namespace(output, "w15", WORD_2012_NAMESPACE, wrote);
    if let Some(value) = value {
        output.push_str(" w15:val=\"");
        output.push_str(value);
        output.push('"');
    }
    output.push_str("/>");
}

fn write_namespace(output: &mut String, prefix: &str, namespace: &str, wrote: &mut bool) {
    if !*wrote {
        output.push_str(" xmlns:");
        output.push_str(prefix);
        output.push_str("=\"");
        output.push_str(namespace);
        output.push('"');
        *wrote = true;
    }
}
