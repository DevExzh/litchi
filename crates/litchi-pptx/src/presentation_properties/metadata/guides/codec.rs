//! Typed, bounded `PowerPoint` 2013 extended presentation guides.

use super::super::{
    BoundedString, Limits, check_string, checked_output_len, escaped_xml_len, process_ooxml,
};
use super::model::{Color, ColorKind, Guide, Guides, List, Orientation};
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion};
use std::collections::HashSet;

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const SLIDE_GUIDES_URI: &str = "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}";
const NOTES_GUIDES_URI: &str = "{2D200454-40CA-4A62-9FC3-DE9A4176ACB9}";

impl Guides {
    /// Parse guide extensions from a complete `p:presentation` document.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_xml(xml: &[u8], limits: Limits) -> Result<Self> {
        if xml.len() > limits.max_input_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::InputBytes,
                xml.len(),
                limits.max_input_bytes(),
            ));
        }
        let processed = process_ooxml(xml, limits)?;
        if processed.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                processed.len(),
                limits.max_output_bytes(),
            ));
        }
        parse_presentation(&parse_dom(processed.as_ref(), limits)?, limits)
    }

    /// Serialize the guide entries as a complete `p:extLst` fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml(&self, strict: bool, limits: Limits) -> Result<String> {
        validate(self, limits)?;
        if self.slide.is_none() && self.notes.is_none() {
            return Ok(String::new());
        }
        let mut xml = BoundedString::new(limits, "serialized presentation guides");
        xml.push_str("<p:extLst>")?;
        if let Some(list) = &self.slide {
            write_list_extension(
                &mut xml,
                SLIDE_GUIDES_URI,
                "sldGuideLst",
                list,
                strict,
                limits,
                None,
            )?;
        }
        if let Some(list) = &self.notes {
            write_list_extension(
                &mut xml,
                NOTES_GUIDES_URI,
                "notesGuideLst",
                list,
                strict,
                limits,
                None,
            )?;
        }
        xml.push_str("</p:extLst>")?;
        Ok(xml.finish())
    }
}

/// Validate a programmatically staged guide value before package publication.
pub(crate) fn validate_value(value: &Guides, limits: Limits) -> Result<()> {
    validate(value, limits)
}

/// Replace only the known guide extension entries in one presentation source.
///
/// The source is preprocessed through the shared MCE baseline before a changed
/// edit is written. The no-op transaction path deliberately skips this helper,
/// so an untouched source is returned byte-for-byte unchanged.
pub(crate) fn rewrite_source(source: &[u8], value: &Guides, limits: Limits) -> Result<Vec<u8>> {
    if source.len() > limits.max_input_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::InputBytes,
            source.len(),
            limits.max_input_bytes(),
        ));
    }
    validate(value, limits)?;
    let processed = process_ooxml(source, limits)?;
    if processed.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            processed.len(),
            limits.max_output_bytes(),
        ));
    }
    let source = processed.as_ref();
    let layout = scan_source(source, limits)?;
    // Count each replacement with the same bounded writer used for the real
    // serialization.  This dry run has no output buffer, so the final source
    // delta can be admitted before either guide entry is allocated.
    let entry_lengths =
        target_entry_lengths(value, layout.strict, layout.presentation_namespace, limits)?;

    let Some(ext_list) = layout.ext_list.as_ref() else {
        if entry_lengths.iter().all(Option::is_none) {
            if source.len() > limits.max_output_bytes() {
                return Err(limit(
                    crate::opened::CodecLimitKind::OutputBytes,
                    source.len(),
                    limits.max_output_bytes(),
                ));
            }
            return Ok(source.to_vec());
        }
        let fragment_len = guide_fragment_len(&entry_lengths, layout.presentation_namespace)?;
        if layout.root.empty {
            let replacement_len = expanded_empty_len(source, &layout.root, fragment_len)?;
            checked_output_len(
                source.len(),
                layout.root.end - layout.root.start,
                replacement_len,
                limits,
                "patched presentation guides",
            )?;
        } else {
            checked_output_len(
                source.len(),
                0,
                fragment_len,
                limits,
                "patched presentation guides",
            )?;
        }
        let entries = target_entries(value, layout.strict, layout.presentation_namespace, limits)?;
        let mut fragment = BoundedString::new(limits, "presentation-guide patch fragment");
        fragment.push_str("<p:extLst xmlns:p=\"")?;
        fragment.push_escaped_xml(layout.presentation_namespace)?;
        fragment.push_str("\">")?;
        for entry in entries.iter().filter_map(Option::as_deref) {
            fragment.push_str(entry)?;
        }
        fragment.push_str("</p:extLst>")?;
        let fragment = fragment.finish();
        return if layout.root.empty {
            let replacement = expand_empty(source, &layout.root, &fragment, limits)?;
            replace_spans(
                source,
                &[(layout.root.start, layout.root.end, replacement)],
                limits,
            )
        } else {
            let root_close = layout
                .root_close
                .ok_or_else(|| invalid("presentation root is missing its end tag"))?;
            insert_bytes(source, root_close, fragment.as_bytes(), limits)
        };
    };

    // Keep replacement ranges and lengths while calculating the exact
    // combined output size.  The serialized entry strings and their owned
    // replacement buffers are not created until the final-size admission
    // succeeds.
    let mut ranges: [(usize, usize, usize, usize); 2] = [(0, 0, 0, 0), (0, 0, 0, 0)];
    let mut plan_count = 0usize;
    for (index, _) in [Target::Slide, Target::Notes].into_iter().enumerate() {
        let replacement_len = entry_lengths[index].unwrap_or_default();
        if let Some(span) = layout.targets[index].as_ref() {
            ranges[plan_count] = (span.start, span.end, replacement_len, index);
            plan_count += 1;
        } else if replacement_len != 0 {
            if ext_list.empty {
                let body_len = entry_lengths.iter().filter_map(|entry| *entry).try_fold(
                    0usize,
                    |length, entry| {
                        length.checked_add(entry).ok_or(Error::CapacityOverflow {
                            resource: "patched presentation guides",
                        })
                    },
                )?;
                let replacement_len = expanded_empty_len(source, ext_list, body_len)?;
                checked_output_len(
                    source.len(),
                    ext_list.end - ext_list.start,
                    replacement_len,
                    limits,
                    "patched presentation guides",
                )?;
                let entries =
                    target_entries(value, layout.strict, layout.presentation_namespace, limits)?;
                let mut body = BoundedString::new(limits, "presentation-guide patch fragment");
                for entry in entries.iter().filter_map(Option::as_deref) {
                    body.push_str(entry)?;
                }
                let body = body.finish();
                let replacement = expand_empty_bytes(source, ext_list, body.as_bytes(), limits)?;
                return replace_spans(
                    source,
                    &[(ext_list.start, ext_list.end, replacement)],
                    limits,
                );
            }
            ranges[plan_count] = (
                ext_list.close_start,
                ext_list.close_start,
                replacement_len,
                index,
            );
            plan_count += 1;
        }
    }

    ranges[..plan_count].sort_unstable_by_key(|(start, end, _, target)| (*start, *end, *target));
    let mut removed = 0usize;
    let mut added = 0usize;
    for pair in ranges[..plan_count].windows(2) {
        if pair[0].1 > pair[1].0 {
            return Err(invalid("overlapping presentation-guide patch ranges"));
        }
    }
    for (start, end, replacement_len, _) in ranges[..plan_count].iter() {
        if *start > *end || *end > source.len() {
            return Err(invalid("presentation-guide patch range is outside source"));
        }
        removed = removed
            .checked_add(*end - *start)
            .ok_or(Error::CapacityOverflow {
                resource: "patched presentation guides",
            })?;
        added = added
            .checked_add(*replacement_len)
            .ok_or(Error::CapacityOverflow {
                resource: "patched presentation guides",
            })?;
    }
    checked_output_len(
        source.len(),
        removed,
        added,
        limits,
        "patched presentation guides",
    )?;

    let entries = target_entries(value, layout.strict, layout.presentation_namespace, limits)?;
    let mut plans: [(usize, usize, &[u8]); 2] = [(0, 0, &[]), (0, 0, &[])];
    for (index, (start, end, _, target_index)) in ranges[..plan_count].iter().enumerate() {
        let replacement = entries[*target_index]
            .as_deref()
            .unwrap_or_default()
            .as_bytes();
        plans[index] = (*start, *end, replacement);
    }
    let mut replacements = Vec::with_capacity(plan_count);
    for (start, end, replacement) in plans[..plan_count].iter() {
        replacements.push((*start, *end, replacement.to_vec()));
    }
    replace_spans(source, &replacements, limits)
}

#[derive(Clone, Copy)]
enum Target {
    Slide,
    Notes,
}

#[derive(Clone)]
struct SourceSpan {
    start: usize,
    end: usize,
    close_start: usize,
    empty: bool,
    qname: String,
}

struct SourceFrame {
    start: usize,
    local: String,
    qname: String,
    ext_list: bool,
    target: Option<Target>,
}

struct SourceLayout {
    root: SourceSpan,
    root_close: Option<usize>,
    ext_list: Option<SourceSpan>,
    targets: [Option<SourceSpan>; 2],
    presentation_namespace: &'static str,
    strict: bool,
}

fn scan_source(source: &[u8], limits: Limits) -> Result<SourceLayout> {
    let mut reader = NsReader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    let mut stack: Vec<SourceFrame> = Vec::new();
    let mut root = None;
    let mut root_close = None;
    let mut ext_list = None;
    let mut targets = [None, None];
    let mut root_namespace = None;
    let mut nodes = 0usize;

    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("presentation-guide XML offset overflow"))?;
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if nodes > limits.max_nodes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes(),
                    ));
                }
                if stack.len() >= limits.max_depth() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth(),
                    ));
                }
                check_source_attributes(&element, decoder, limits)?;
                if element.name().as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::NameBytes,
                        element.name().as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let local = local_name(element.local_name().as_ref())?;
                let qname = std::str::from_utf8(element.name().as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                let is_presentation = presentation_namespace(&namespace).is_some();
                if stack.is_empty() {
                    if root.is_some() || !is_presentation || local != "presentation" {
                        return Err(invalid("expected one PresentationML presentation root"));
                    }
                    root_namespace = presentation_namespace(&namespace);
                }
                let is_ext_list = stack.len() == 1 && is_presentation && local == "extLst";
                if is_ext_list && ext_list.is_some() {
                    return Err(invalid("duplicate presentation extLst"));
                }
                let target = if stack.last().is_some_and(|frame| frame.ext_list)
                    && is_presentation
                    && local == "ext"
                {
                    target_from_uri(element_uri(&element, decoder, limits)?)
                } else {
                    None
                };
                stack.push(SourceFrame {
                    start,
                    local,
                    qname,
                    ext_list: is_ext_list,
                    target,
                });
            },
            Event::Empty(element) => {
                check_empty_depth(stack.len(), limits)?;
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if nodes > limits.max_nodes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes(),
                    ));
                }
                check_source_attributes(&element, decoder, limits)?;
                if element.name().as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::NameBytes,
                        element.name().as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let local = local_name(element.local_name().as_ref())?;
                let qname = std::str::from_utf8(element.name().as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                let is_presentation = presentation_namespace(&namespace).is_some();
                if stack.is_empty() {
                    if root.is_some() || !is_presentation || local != "presentation" {
                        return Err(invalid("expected one PresentationML presentation root"));
                    }
                    root_namespace = presentation_namespace(&namespace);
                    root = Some(SourceSpan {
                        start,
                        end: origin.offset(reader.buffer_position()).ok_or_else(|| {
                            invalid("presentation-guide XML position exceeds usize")
                        })?,
                        close_start: start,
                        empty: true,
                        qname,
                    });
                } else if stack.len() == 1 && is_presentation && local == "extLst" {
                    if ext_list.is_some() {
                        return Err(invalid("duplicate presentation extLst"));
                    }
                    ext_list = Some(SourceSpan {
                        start,
                        end: origin.offset(reader.buffer_position()).ok_or_else(|| {
                            invalid("presentation-guide XML position exceeds usize")
                        })?,
                        close_start: start,
                        empty: true,
                        qname,
                    });
                } else if stack.len() == 2
                    && stack.last().is_some_and(|frame| frame.ext_list)
                    && is_presentation
                    && local == "ext"
                    && let Some(target) = target_from_uri(element_uri(&element, decoder, limits)?)
                {
                    let index = target_index(target);
                    if targets[index].is_some() {
                        return Err(invalid("duplicate extended-guide extension"));
                    }
                    targets[index] = Some(SourceSpan {
                        start,
                        end: origin.offset(reader.buffer_position()).ok_or_else(|| {
                            invalid("presentation-guide XML position exceeds usize")
                        })?,
                        close_start: start,
                        empty: true,
                        qname,
                    });
                }
            },
            Event::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected presentation-guide closing element"))?;
                let local = local_name(element.local_name().as_ref())?;
                if frame.local != local {
                    return Err(invalid("mismatched presentation-guide closing element"));
                }
                let end = origin
                    .offset(reader.buffer_position())
                    .ok_or_else(|| invalid("presentation-guide XML offset overflow"))?;
                let span = SourceSpan {
                    start: frame.start,
                    end,
                    close_start: start,
                    empty: false,
                    qname: frame.qname,
                };
                if stack.is_empty() {
                    if root.is_some() {
                        return Err(invalid("multiple PresentationML presentation roots"));
                    }
                    root_close = Some(start);
                    root = Some(span);
                } else if frame.ext_list {
                    if ext_list.is_some() {
                        return Err(invalid("duplicate presentation extLst"));
                    }
                    ext_list = Some(span);
                } else if let Some(target) = frame.target {
                    let index = target_index(target);
                    if targets[index].is_some() {
                        return Err(invalid("duplicate extended-guide extension"));
                    }
                    targets[index] = Some(span);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Text(_) | Event::CData(_) | Event::Comment(_) => {},
            Event::GeneralRef(_) => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation-guide XML"));
    }
    let root = root.ok_or_else(|| invalid("missing presentation root"))?;
    let presentation_namespace =
        root_namespace.ok_or_else(|| invalid("missing PresentationML presentation namespace"))?;
    Ok(SourceLayout {
        strict: presentation_namespace == PS,
        root,
        root_close,
        ext_list,
        targets,
        presentation_namespace,
    })
}

fn local_name(name: &[u8]) -> Result<String> {
    let name = std::str::from_utf8(name).map_err(xml_error)?;
    Ok(name
        .rsplit_once(':')
        .map_or(name, |(_, local)| local)
        .to_owned())
}

fn presentation_namespace(namespace: &ResolveResult<'_>) -> Option<&'static str> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == P.as_bytes() => Some(P),
        ResolveResult::Bound(Namespace(value)) if *value == PS.as_bytes() => Some(PS),
        _ => None,
    }
}

fn element_uri(
    element: &BytesStart<'_>,
    decoder: Decoder,
    limits: Limits,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"uri" {
            if value.is_some() {
                return Err(invalid("duplicate presentation-guide extension URI"));
            }
            if attribute.value.as_ref().len() > limits.max_string_bytes() {
                return Err(limit(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                    attribute.value.as_ref().len(),
                    limits.max_string_bytes(),
                ));
            }
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?;
            if decoded.len() > limits.max_string_bytes() {
                return Err(limit(
                    crate::opened::CodecLimitKind::StringBytes,
                    decoded.len(),
                    limits.max_string_bytes(),
                ));
            }
            value = Some(decoded.into_owned());
        }
    }
    Ok(value)
}

fn check_source_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    limits: Limits,
) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        count += 1;
        if count > limits.max_attributes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributes,
                count,
                limits.max_attributes(),
            ));
        }
        if attribute.key.as_ref().len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributeBytes,
                attribute.key.as_ref().len(),
                limits.max_string_bytes(),
            ));
        }
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributeBytes,
                attribute.value.as_ref().len(),
                limits.max_string_bytes(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        if value.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::StringBytes,
                value.len(),
                limits.max_string_bytes(),
            ));
        }
    }
    Ok(())
}

fn target_from_uri(uri: Option<String>) -> Option<Target> {
    match uri.as_deref() {
        Some(SLIDE_GUIDES_URI) => Some(Target::Slide),
        Some(NOTES_GUIDES_URI) => Some(Target::Notes),
        _ => None,
    }
}

fn target_index(target: Target) -> usize {
    match target {
        Target::Slide => 0,
        Target::Notes => 1,
    }
}

fn target_entries(
    value: &Guides,
    strict: bool,
    presentation_namespace: &str,
    limits: Limits,
) -> Result<[Option<String>; 2]> {
    Ok([
        value
            .slide
            .as_ref()
            .map(|list| {
                extension_xml(
                    SLIDE_GUIDES_URI,
                    "sldGuideLst",
                    list,
                    strict,
                    presentation_namespace,
                    limits,
                )
            })
            .transpose()?,
        value
            .notes
            .as_ref()
            .map(|list| {
                extension_xml(
                    NOTES_GUIDES_URI,
                    "notesGuideLst",
                    list,
                    strict,
                    presentation_namespace,
                    limits,
                )
            })
            .transpose()?,
    ])
}

fn target_entry_lengths(
    value: &Guides,
    strict: bool,
    presentation_namespace: &str,
    limits: Limits,
) -> Result<[Option<usize>; 2]> {
    Ok([
        value
            .slide
            .as_ref()
            .map(|list| {
                extension_len(
                    SLIDE_GUIDES_URI,
                    "sldGuideLst",
                    list,
                    strict,
                    presentation_namespace,
                    limits,
                )
            })
            .transpose()?,
        value
            .notes
            .as_ref()
            .map(|list| {
                extension_len(
                    NOTES_GUIDES_URI,
                    "notesGuideLst",
                    list,
                    strict,
                    presentation_namespace,
                    limits,
                )
            })
            .transpose()?,
    ])
}

fn extension_xml(
    uri: &str,
    local: &str,
    list: &List,
    strict: bool,
    presentation_namespace: &str,
    limits: Limits,
) -> Result<String> {
    let mut xml = BoundedString::new(limits, "guide extension output");
    write_list_extension(
        &mut xml,
        uri,
        local,
        list,
        strict,
        limits,
        Some(presentation_namespace),
    )?;
    Ok(xml.finish())
}

fn extension_len(
    uri: &str,
    local: &str,
    list: &List,
    strict: bool,
    presentation_namespace: &str,
    limits: Limits,
) -> Result<usize> {
    let mut output = LengthCounter::new(limits, "guide extension output");
    write_list_extension(
        &mut output,
        uri,
        local,
        list,
        strict,
        limits,
        Some(presentation_namespace),
    )?;
    Ok(output.finish())
}

fn expand_empty(source: &[u8], span: &SourceSpan, body: &str, limits: Limits) -> Result<Vec<u8>> {
    let replacement = expand_empty_bytes(source, span, body.as_bytes(), limits)?;
    Ok(replacement)
}

fn guide_fragment_len(
    entry_lengths: &[Option<usize>; 2],
    presentation_namespace: &str,
) -> Result<usize> {
    let mut length = b"<p:extLst xmlns:p=\"".len();
    length = length
        .checked_add(escaped_xml_len(presentation_namespace)?)
        .and_then(|value| value.checked_add(b"\">".len()))
        .ok_or(Error::CapacityOverflow {
            resource: "patched presentation guides",
        })?;
    for entry in entry_lengths.iter().filter_map(|entry| *entry) {
        length = length.checked_add(entry).ok_or(Error::CapacityOverflow {
            resource: "patched presentation guides",
        })?;
    }
    length
        .checked_add(b"</p:extLst>".len())
        .ok_or(Error::CapacityOverflow {
            resource: "patched presentation guides",
        })
}

fn expanded_empty_len(source: &[u8], span: &SourceSpan, body_len: usize) -> Result<usize> {
    let token = source[span.start..span.end].trim_ascii_end();
    let open = token
        .strip_suffix(b"/>")
        .ok_or_else(|| invalid("presentation-guide empty element has no self-close"))?;
    open.len()
        .checked_add(1)
        .and_then(|value| value.checked_add(body_len))
        .and_then(|value| value.checked_add(span.qname.len()))
        .and_then(|value| value.checked_add(3))
        .ok_or(Error::CapacityOverflow {
            resource: "patched presentation guides",
        })
}

fn expand_empty_bytes(
    source: &[u8],
    span: &SourceSpan,
    body: &[u8],
    limits: Limits,
) -> Result<Vec<u8>> {
    let token = source[span.start..span.end].trim_ascii_end();
    let open = token
        .strip_suffix(b"/>")
        .ok_or_else(|| invalid("presentation-guide empty element has no self-close"))?;
    let replacement_len = expanded_empty_len(source, span, body.len())?;
    if replacement_len > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            replacement_len,
            limits.max_output_bytes(),
        ));
    }
    let mut replacement = Vec::new();
    replacement
        .try_reserve(replacement_len)
        .map_err(|source| Error::Allocation {
            resource: "patched presentation guides",
            source,
        })?;
    replacement.extend_from_slice(open);
    replacement.push(b'>');
    replacement.extend_from_slice(body);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(span.qname.as_bytes());
    replacement.push(b'>');
    debug_assert_eq!(replacement.len(), replacement_len);
    Ok(replacement)
}

fn replace_spans(
    source: &[u8],
    spans: &[(usize, usize, Vec<u8>)],
    limits: Limits,
) -> Result<Vec<u8>> {
    let mut ordered = spans.to_vec();
    ordered.sort_by_key(|(start, _, _)| *start);
    for pair in ordered.windows(2) {
        if pair[0].1 > pair[1].0 {
            return Err(invalid("overlapping presentation-guide patch ranges"));
        }
    }
    let mut removed = 0usize;
    let mut added = 0usize;
    for (start, end, replacement) in &ordered {
        if *start > *end || *end > source.len() {
            return Err(invalid("presentation-guide patch range is outside source"));
        }
        removed = removed
            .checked_add(*end - *start)
            .ok_or(Error::CapacityOverflow {
                resource: "patched presentation guides",
            })?;
        added = added
            .checked_add(replacement.len())
            .ok_or(Error::CapacityOverflow {
                resource: "patched presentation guides",
            })?;
    }
    let output_len = checked_output_len(
        source.len(),
        removed,
        added,
        limits,
        "patched presentation guides",
    )?;
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|source| Error::Allocation {
            resource: "patched presentation guides",
            source,
        })?;
    let mut cursor = 0usize;
    for (start, end, replacement) in ordered {
        output.extend_from_slice(&source[cursor..start]);
        output.extend_from_slice(&replacement);
        cursor = end;
    }
    output.extend_from_slice(&source[cursor..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn insert_bytes(
    source: &[u8],
    offset: usize,
    replacement: &[u8],
    limits: Limits,
) -> Result<Vec<u8>> {
    if offset > source.len() {
        return Err(invalid(
            "presentation-guide insertion offset is out of bounds",
        ));
    }
    let output_len =
        source
            .len()
            .checked_add(replacement.len())
            .ok_or(Error::CapacityOverflow {
                resource: "patched presentation guides",
            })?;
    if output_len > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            output_len,
            limits.max_output_bytes(),
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|source| Error::Allocation {
            resource: "patched presentation guides",
            source,
        })?;
    output.extend_from_slice(&source[..offset]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[offset..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

#[derive(Clone)]
struct Attr {
    qname: String,
    namespace: String,
    local: String,
    value: String,
}

#[derive(Clone)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}

#[derive(Clone)]
struct Node {
    qname: String,
    namespace: String,
    local: String,
    attributes: Vec<Attr>,
    bindings: Vec<(String, String)>,
    content: Vec<Content>,
}

fn parse_dom(xml: &[u8], limits: Limits) -> Result<Node> {
    if xml.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            xml.len(),
            limits.max_output_bytes(),
        ));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut count = 0usize;
    loop {
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                count = count.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if count > limits.max_nodes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        count,
                        limits.max_nodes(),
                    ));
                }
                if stack.len() >= limits.max_depth() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth(),
                    ));
                }
                stack.push(make_node(&element, decoder, &stack, limits)?);
            },
            Ok(Event::Empty(element)) => {
                check_empty_depth(stack.len(), limits)?;
                count = count.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if count > limits.max_nodes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        count,
                        limits.max_nodes(),
                    ));
                }
                let node = make_node(&element, decoder, &stack, limits)?;
                attach(&mut stack, &mut root, node)?;
            },
            Ok(Event::End(_)) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                attach(&mut stack, &mut root, node)?;
            },
            Ok(Event::Text(text)) => {
                if text.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::TextBytes,
                        text.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.decode().map_err(xml_error)?;
                if text.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::StringBytes,
                        text.len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.into_owned();
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::Text(text));
                } else if !text.trim().is_empty() {
                    return Err(invalid("text outside presentation root"));
                }
            },
            Ok(Event::CData(text)) => {
                if text.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::TextBytes,
                        text.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.decode().map_err(xml_error)?;
                if text.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::StringBytes,
                        text.len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.into_owned();
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::CData(text));
                } else {
                    return Err(invalid("CDATA outside presentation root"));
                }
            },
            Ok(Event::Comment(text)) => {
                if text.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::TextBytes,
                        text.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.decode().map_err(xml_error)?;
                if text.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::StringBytes,
                        text.len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.into_owned();
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::Comment(text));
                }
            },
            Ok(Event::GeneralRef(reference)) => {
                if reference.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::ReferenceBytes,
                        reference.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = litchi_ooxml_common::xml::decode_xml_reference(&reference)?;
                if text.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::StringBytes,
                        text.len(),
                        limits.max_string_bytes(),
                    ));
                }
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::Text(text));
                } else {
                    return Err(invalid("entity outside presentation root"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(error) => return Err(xml_error(error)),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation-guide XML"));
    }
    root.ok_or_else(|| invalid("missing presentation root"))
}

fn make_node(
    element: &BytesStart<'_>,
    decoder: Decoder,
    stack: &[Node],
    limits: Limits,
) -> Result<Node> {
    let raw_name = element.name();
    let qname = std::str::from_utf8(raw_name.as_ref()).map_err(xml_error)?;
    bounded_string(qname, limits)?;
    let qname = qname.to_owned();
    let mut bindings = stack
        .last()
        .map(|node| node.bindings.clone())
        .unwrap_or_default();
    let mut raw = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if raw.len() >= limits.max_attributes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributes,
                raw.len() + 1,
                limits.max_attributes(),
            ));
        }
        let name = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        bounded_string(name, limits)?;
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributeBytes,
                attribute.value.as_ref().len(),
                limits.max_string_bytes(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        if value.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::StringBytes,
                value.len(),
                limits.max_string_bytes(),
            ));
        }
        raw.push((name.to_owned(), value.into_owned()));
    }
    for (name, value) in &raw {
        if name == "xmlns" || name.starts_with("xmlns:") {
            let prefix = name.strip_prefix("xmlns:").unwrap_or("").to_owned();
            if let Some(binding) = bindings.iter_mut().find(|binding| binding.0 == prefix) {
                binding.1.clone_from(value);
            } else {
                bindings.push((prefix, value.clone()));
            }
        }
    }
    let (prefix, local) = split_qname(&qname)?;
    let namespace = resolve(&bindings, prefix)?;
    let local = local.to_owned();
    let mut attributes = Vec::new();
    for (name, value) in raw {
        if name == "xmlns" || name.starts_with("xmlns:") {
            continue;
        }
        let (prefix, local) = split_qname(&name)?;
        let namespace = if prefix.is_empty() {
            String::new()
        } else {
            resolve(&bindings, prefix)?
        };
        let local = local.to_owned();
        attributes.push(Attr {
            qname: name,
            namespace,
            local,
            value,
        });
    }
    Ok(Node {
        qname,
        namespace,
        local,
        attributes,
        bindings,
        content: Vec::new(),
    })
}

fn attach(stack: &mut [Node], root: &mut Option<Node>, node: Node) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.content.push(Content::Node(node));
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn parse_presentation(root: &Node, limits: Limits) -> Result<Guides> {
    expect(root, &[P, PS], "presentation")?;
    let mut root_ext = None;
    for child in children(root)? {
        if is(child, &[P, PS], "extLst") && root_ext.replace(child).is_some() {
            return Err(invalid("duplicate presentation extLst"));
        }
    }
    let Some(root_ext) = root_ext else {
        return Ok(Guides::default());
    };
    let mut value = Guides::default();
    for extension in children(root_ext)? {
        expect(extension, &[P, PS], "ext")?;
        let uri = required_attr(extension, "uri")?;
        let target = match uri.as_str() {
            SLIDE_GUIDES_URI => Some((&mut value.slide, "sldGuideLst")),
            NOTES_GUIDES_URI => Some((&mut value.notes, "notesGuideLst")),
            _ => None,
        };
        if let Some((slot, local)) = target {
            if slot.is_some() {
                return Err(invalid(format!("duplicate {local} extension")));
            }
            only_unqualified(extension, &["uri"])?;
            let payload = children(extension)?;
            if payload.len() != 1 {
                return Err(invalid(format!("{local} extension requires one payload")));
            }
            expect(payload[0], &[P15], local)?;
            *slot = Some(parse_list(payload[0], limits)?);
        }
    }
    Ok(value)
}

fn parse_list(node: &Node, limits: Limits) -> Result<List> {
    no_attributes(node)?;
    let content = children(node)?;
    let mut guides = Vec::new();
    let mut extension_xml = None;
    let mut ids = HashSet::new();
    for child in content {
        if is(child, &[P15], "guide") {
            if extension_xml.is_some() {
                return Err(invalid("guide appears after guide-list extLst"));
            }
            if guides.len() >= limits.max_guides() {
                return Err(limit(
                    crate::opened::CodecLimitKind::Guides,
                    guides.len() + 1,
                    limits.max_guides(),
                ));
            }
            let guide = parse_guide(child, limits)?;
            if !ids.insert(guide.id) {
                return Err(invalid("duplicate extended guide ID"));
            }
            guides.push(guide);
        } else if is(child, &[P, PS], "extLst") {
            if extension_xml.is_some() {
                return Err(invalid("duplicate guide-list extLst"));
            }
            validate_extension_list(child, limits)?;
            extension_xml = Some(node_xml(child, false, limits)?);
        } else {
            return Err(invalid("unexpected extended guide-list child"));
        }
    }
    Ok(List {
        guides,
        extension_xml,
    })
}

fn parse_guide(node: &Node, limits: Limits) -> Result<Guide> {
    let id = required_attr(node, "id")?
        .parse::<u32>()
        .map_err(|_err| invalid("invalid extended guide ID"))?;
    let name = optional_attr(node, "name")?;
    if let Some(name) = &name {
        bounded_string(name, limits)?;
    }
    let orientation = optional_attr(node, "orient")?
        .map(|value| match value.as_str() {
            "horz" => Ok(Orientation::Horizontal),
            "vert" => Ok(Orientation::Vertical),
            _ => Err(invalid("invalid extended guide orientation")),
        })
        .transpose()?;
    let position = optional_attr(node, "pos")?
        .map(|value| {
            value
                .parse::<i32>()
                .map_err(|_err| invalid("invalid extended guide position"))
        })
        .transpose()?;
    let user_drawn = optional_attr(node, "userDrawn")?
        .map(|value| parse_bool(&value, "userDrawn"))
        .transpose()?;
    only_unqualified(node, &["id", "name", "orient", "pos", "userDrawn"])?;
    let content = children(node)?;
    if content.is_empty()
        || content.len() > 2
        || !is(content[0], &[P15], "clr")
        || (content.len() == 2 && !is(content[1], &[P, PS], "extLst"))
    {
        return Err(invalid("guide requires clr followed by optional extLst"));
    }
    let color = parse_color(content[0], limits)?;
    let extension_xml = if content.len() == 2 {
        validate_extension_list(content[1], limits)?;
        Some(node_xml(content[1], false, limits)?)
    } else {
        None
    };
    Ok(Guide {
        id,
        name,
        orientation,
        position,
        user_drawn,
        color,
        extension_xml,
    })
}

fn parse_color(node: &Node, limits: Limits) -> Result<Color> {
    no_attributes(node)?;
    let colors = children(node)?;
    if colors.len() != 1 {
        return Err(invalid("extended guide clr requires one DrawingML color"));
    }
    let color = colors[0];
    let kind = match (color.namespace.as_str(), color.local.as_str()) {
        (A | AS, "scrgbClr") => ColorKind::ScRgb,
        (A | AS, "srgbClr") => ColorKind::Srgb,
        (A | AS, "hslClr") => ColorKind::Hsl,
        (A | AS, "sysClr") => ColorKind::System,
        (A | AS, "schemeClr") => ColorKind::Scheme,
        (A | AS, "prstClr") => ColorKind::Preset,
        _ => return Err(invalid("invalid extended guide DrawingML color")),
    };
    Ok(Color {
        kind,
        xml: node_xml(color, false, limits)?,
    })
}

fn validate_extension_list(node: &Node, limits: Limits) -> Result<()> {
    expect(node, &[P, PS], "extLst")?;
    no_attributes(node)?;
    let extensions = children(node)?;
    if extensions.len() > limits.max_extensions() {
        return Err(limit(
            crate::opened::CodecLimitKind::Extensions,
            extensions.len(),
            limits.max_extensions(),
        ));
    }
    for extension in extensions {
        expect(extension, &[P, PS], "ext")?;
        let uri = required_attr(extension, "uri")?;
        if uri.is_empty() {
            return Err(invalid("guide extension URI is empty"));
        }
        bounded_string(&uri, limits)?;
        only_unqualified(extension, &["uri"])?;
    }
    Ok(())
}

fn validate(value: &Guides, limits: Limits) -> Result<()> {
    for list in value.slide.iter().chain(value.notes.iter()) {
        if list.guides.len() > limits.max_guides() {
            return Err(limit(
                crate::opened::CodecLimitKind::Guides,
                list.guides.len(),
                limits.max_guides(),
            ));
        }
        let mut ids = HashSet::new();
        for guide in &list.guides {
            if !ids.insert(guide.id) {
                return Err(invalid("duplicate extended guide ID"));
            }
            if let Some(name) = &guide.name {
                bounded_string(name, limits)?;
            }
            validate_color(&guide.color, limits)?;
            if let Some(extension) = &guide.extension_xml {
                validate_opaque_extension(extension, limits)?;
            }
        }
        if let Some(extension) = &list.extension_xml {
            validate_opaque_extension(extension, limits)?;
        }
    }
    Ok(())
}

fn validate_color(color: &Color, limits: Limits) -> Result<()> {
    if color.xml.len() > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::StringBytes,
            color.xml.len(),
            limits.max_string_bytes(),
        ));
    }
    let node = parse_dom(&color.xml, limits)?;
    let parsed = parse_color_node(&node)?;
    if parsed != color.kind {
        return Err(invalid("extended guide color kind does not match XML"));
    }
    Ok(())
}

fn parse_color_node(node: &Node) -> Result<ColorKind> {
    match (node.namespace.as_str(), node.local.as_str()) {
        (A | AS, "scrgbClr") => Ok(ColorKind::ScRgb),
        (A | AS, "srgbClr") => Ok(ColorKind::Srgb),
        (A | AS, "hslClr") => Ok(ColorKind::Hsl),
        (A | AS, "sysClr") => Ok(ColorKind::System),
        (A | AS, "schemeClr") => Ok(ColorKind::Scheme),
        (A | AS, "prstClr") => Ok(ColorKind::Preset),
        _ => Err(invalid("invalid extended guide DrawingML color")),
    }
}

fn validate_opaque_extension(xml: &[u8], limits: Limits) -> Result<()> {
    if xml.len() > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::ExtensionBytes,
            xml.len(),
            limits.max_string_bytes(),
        ));
    }
    validate_extension_list(&parse_dom(xml, limits)?, limits)
}

trait GuideWriter {
    fn push_str(&mut self, value: &str) -> Result<()>;
    fn push(&mut self, value: char) -> Result<()>;
    fn push_escaped_xml(&mut self, value: &str) -> Result<()>;
    fn write_fmt(&mut self, arguments: std::fmt::Arguments<'_>) -> Result<()>;
}

impl GuideWriter for BoundedString {
    fn push_str(&mut self, value: &str) -> Result<()> {
        BoundedString::push_str(self, value)
    }

    fn push(&mut self, value: char) -> Result<()> {
        BoundedString::push(self, value)
    }

    fn push_escaped_xml(&mut self, value: &str) -> Result<()> {
        BoundedString::push_escaped_xml(self, value)
    }

    fn write_fmt(&mut self, arguments: std::fmt::Arguments<'_>) -> Result<()> {
        BoundedString::write_fmt(self, arguments)
    }
}

struct LengthCounter {
    length: usize,
    limit: usize,
    resource: &'static str,
    failure: Option<Error>,
}

impl LengthCounter {
    fn new(limits: Limits, resource: &'static str) -> Self {
        Self {
            length: 0,
            limit: limits.max_output_bytes(),
            resource,
            failure: None,
        }
    }

    fn reserve_for(&mut self, additional: usize) -> Result<()> {
        let required = self
            .length
            .checked_add(additional)
            .ok_or(Error::CapacityOverflow {
                resource: self.resource,
            })?;
        if required > self.limit {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                required,
                self.limit,
            ));
        }
        self.length = required;
        Ok(())
    }

    fn finish(self) -> usize {
        self.length
    }
}

impl std::fmt::Write for LengthCounter {
    fn write_str(&mut self, value: &str) -> std::fmt::Result {
        match self.reserve_for(value.len()) {
            Ok(()) => Ok(()),
            Err(error) => {
                self.failure = Some(error);
                Err(std::fmt::Error)
            },
        }
    }
}

impl GuideWriter for LengthCounter {
    fn push_str(&mut self, value: &str) -> Result<()> {
        self.reserve_for(value.len())
    }

    fn push(&mut self, value: char) -> Result<()> {
        self.reserve_for(value.len_utf8())
    }

    fn push_escaped_xml(&mut self, value: &str) -> Result<()> {
        self.reserve_for(escaped_xml_len(value)?)
    }

    fn write_fmt(&mut self, arguments: std::fmt::Arguments<'_>) -> Result<()> {
        std::fmt::write(self, arguments).map_err(|_| {
            self.failure
                .take()
                .unwrap_or_else(|| invalid("guide output formatting failed"))
        })
    }
}

fn write_list_extension<W: GuideWriter>(
    xml: &mut W,
    uri: &str,
    local: &str,
    list: &List,
    strict: bool,
    limits: Limits,
    presentation_namespace: Option<&str>,
) -> Result<()> {
    if let Some(namespace) = presentation_namespace {
        xml.push_str("<p:ext xmlns:p=\"")?;
        xml.push_escaped_xml(namespace)?;
        xml.write_fmt(format_args!(
            "\" uri=\"{uri}\"><p15:{local} xmlns:p15=\"{P15}\">"
        ))?;
    } else {
        xml.write_fmt(format_args!(
            "<p:ext uri=\"{uri}\"><p15:{local} xmlns:p15=\"{P15}\">"
        ))?;
    }
    for guide in &list.guides {
        write_guide(xml, guide, strict, limits)?;
    }
    if let Some(extension) = &list.extension_xml {
        write_opaque(xml, extension, strict, limits)?;
    }
    xml.write_fmt(format_args!("</p15:{local}></p:ext>"))?;
    Ok(())
}

fn write_guide<W: GuideWriter>(
    xml: &mut W,
    guide: &Guide,
    strict: bool,
    limits: Limits,
) -> Result<()> {
    xml.write_fmt(format_args!("<p15:guide id=\"{}\"", guide.id))?;
    if let Some(name) = &guide.name {
        check_string(name, limits, "extended guide name")?;
        xml.push_str(" name=\"")?;
        escape_attribute(xml, name)?;
        xml.push('"')?;
    }
    if let Some(orientation) = guide.orientation {
        xml.push_str(" orient=\"")?;
        xml.push_str(match orientation {
            Orientation::Horizontal => "horz",
            Orientation::Vertical => "vert",
        })?;
        xml.push('"')?;
    }
    if let Some(position) = guide.position {
        xml.write_fmt(format_args!(" pos=\"{position}\""))?;
    }
    if let Some(user_drawn) = guide.user_drawn {
        xml.write_fmt(format_args!(" userDrawn=\"{}\"", i32::from(user_drawn)))?;
    }
    xml.push_str("><p15:clr>")?;
    write_opaque(xml, &guide.color.xml, strict, limits)?;
    xml.push_str("</p15:clr>")?;
    if let Some(extension) = &guide.extension_xml {
        write_opaque(xml, extension, strict, limits)?;
    }
    xml.push_str("</p15:guide>")?;
    Ok(())
}

fn write_opaque<W: GuideWriter>(
    output: &mut W,
    xml: &[u8],
    strict: bool,
    limits: Limits,
) -> Result<()> {
    if xml.len() > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::ExtensionBytes,
            xml.len(),
            limits.max_string_bytes(),
        ));
    }
    let text = std::str::from_utf8(xml).map_err(xml_error)?;
    push_namespace_rewritten(output, text, strict)?;
    Ok(())
}

fn push_namespace_rewritten<W: GuideWriter>(
    output: &mut W,
    value: &str,
    strict: bool,
) -> Result<()> {
    let replacements: &[(&str, &str)] = if strict {
        &[(P, PS), (A, AS)]
    } else {
        &[(PS, P), (AS, A)]
    };
    let mut cursor = 0usize;
    while cursor < value.len() {
        let next = replacements
            .iter()
            .filter_map(|(from, to)| {
                value[cursor..]
                    .find(from)
                    .map(|offset| (cursor + offset, *from, *to))
            })
            .min_by_key(|(offset, _, _)| *offset);
        let Some((offset, from, to)) = next else {
            output.push_str(&value[cursor..])?;
            return Ok(());
        };
        output.push_str(&value[cursor..offset])?;
        output.push_str(to)?;
        cursor = offset + from.len();
    }
    Ok(())
}

fn children(node: &Node) -> Result<Vec<&Node>> {
    let mut children = Vec::new();
    for content in &node.content {
        match content {
            Content::Node(node) => children.push(node),
            Content::Text(text) if text.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed presentation guides")),
        }
    }
    Ok(children)
}

fn expect(node: &Node, namespaces: &[&str], local: &str) -> Result<()> {
    if is(node, namespaces, local) {
        Ok(())
    } else {
        Err(invalid(format!("expected {local}")))
    }
}

fn is(node: &Node, namespaces: &[&str], local: &str) -> bool {
    namespaces.contains(&node.namespace.as_str()) && node.local == local
}

fn optional_attr(node: &Node, local: &str) -> Result<Option<String>> {
    let mut value = None;
    for attribute in &node.attributes {
        if attribute.namespace.is_empty()
            && attribute.local == local
            && value.replace(attribute.value.clone()).is_some()
        {
            return Err(invalid(format!("duplicate attribute '{local}'")));
        }
    }
    Ok(value)
}

fn required_attr(node: &Node, local: &str) -> Result<String> {
    optional_attr(node, local)?.ok_or_else(|| invalid(format!("missing attribute '{local}'")))
}

fn only_unqualified(node: &Node, allowed: &[&str]) -> Result<()> {
    for attribute in &node.attributes {
        if !attribute.namespace.is_empty() || !allowed.contains(&attribute.local.as_str()) {
            return Err(invalid(format!(
                "unexpected attribute '{}'",
                attribute.qname
            )));
        }
    }
    Ok(())
}

fn no_attributes(node: &Node) -> Result<()> {
    only_unqualified(node, &[])
}

fn bounded_string(value: &str, limits: Limits) -> Result<()> {
    if value.len() > limits.max_string_bytes() {
        Err(limit(
            crate::opened::CodecLimitKind::StringBytes,
            value.len(),
            limits.max_string_bytes(),
        ))
    } else {
        Ok(())
    }
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid boolean '{name}'"))),
    }
}

fn node_xml(node: &Node, strict: bool, limits: Limits) -> Result<Vec<u8>> {
    let mut xml =
        BoundedString::with_limit(limits.max_string_bytes(), "presentation guide node XML");
    write_node(&mut xml, node, strict)?;
    Ok(xml.finish().into_bytes())
}

fn write_node(xml: &mut BoundedString, node: &Node, strict: bool) -> Result<()> {
    let (node_prefix, _) = split_qname(&node.qname)?;
    let mut required_prefixes = HashSet::new();
    if !node.namespace.is_empty() {
        required_prefixes.insert(node_prefix.to_owned());
    }
    for attribute in &node.attributes {
        if let Some((prefix, _)) = attribute.qname.split_once(':') {
            required_prefixes.insert(prefix.to_owned());
        }
    }
    xml.push('<')?;
    xml.push_str(&node.qname)?;
    for (prefix, uri) in &node.bindings {
        if !required_prefixes.contains(prefix) {
            continue;
        }
        if prefix.is_empty() {
            xml.push_str(" xmlns=\"")?;
        } else {
            xml.push_str(" xmlns:")?;
            xml.push_str(prefix)?;
            xml.push_str("=\"")?;
        }
        let uri = if strict {
            match uri.as_str() {
                P => PS,
                A => AS,
                _ => uri,
            }
        } else {
            match uri.as_str() {
                PS => P,
                AS => A,
                _ => uri,
            }
        };
        escape_attribute(xml, uri)?;
        xml.push('"')?;
    }
    for attribute in &node.attributes {
        xml.push(' ')?;
        xml.push_str(&attribute.qname)?;
        xml.push_str("=\"")?;
        escape_attribute(xml, &attribute.value)?;
        xml.push('"')?;
    }
    if node.content.is_empty() {
        xml.push_str("/>")?;
        return Ok(());
    }
    xml.push('>')?;
    for content in &node.content {
        match content {
            Content::Node(node) => write_node(xml, node, strict)?,
            Content::Text(text) => escape_text(xml, text)?,
            Content::CData(text) => {
                xml.push_str("<![CDATA[")?;
                xml.push_str(text)?;
                xml.push_str("]]>")?;
            },
            Content::Comment(text) => {
                xml.push_str("<!--")?;
                xml.push_str(text)?;
                xml.push_str("-->")?;
            },
        }
    }
    xml.push_str("</")?;
    xml.push_str(&node.qname)?;
    xml.push('>')?;
    Ok(())
}

fn split_qname(value: &str) -> Result<(&str, &str)> {
    if let Some((prefix, local)) = value.split_once(':') {
        if local.is_empty() || local.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((prefix, local))
    } else {
        Ok(("", value))
    }
}

fn resolve(bindings: &[(String, String)], prefix: &str) -> Result<String> {
    bindings
        .iter()
        .rev()
        .find(|binding| binding.0 == prefix)
        .map(|binding| binding.1.clone())
        .ok_or_else(|| invalid(format!("unbound namespace prefix '{prefix}'")))
}

fn escape_attribute<W: GuideWriter>(output: &mut W, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;")?,
            '<' => output.push_str("&lt;")?,
            '"' => output.push_str("&quot;")?,
            '\r' => output.push_str("&#xD;")?,
            '\n' => output.push_str("&#xA;")?,
            '\t' => output.push_str("&#x9;")?,
            _ => output.push(character)?,
        }
    }
    Ok(())
}

fn escape_text(output: &mut BoundedString, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;")?,
            '<' => output.push_str("&lt;")?,
            '>' => output.push_str("&gt;")?,
            _ => output.push(character)?,
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::opened::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::opened::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

fn check_empty_depth(depth: usize, limits: Limits) -> Result<()> {
    let observed = depth.checked_add(1).unwrap_or(usize::MAX);
    if observed > limits.max_depth() {
        return Err(limit(
            crate::opened::CodecLimitKind::XmlDepth,
            observed,
            limits.max_depth(),
        ));
    }
    Ok(())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn wrap(fragment: &str) -> String {
        format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{fragment}</p:presentation>"#
        )
    }

    #[test]
    fn round_trips_slide_notes_and_unknown_extensions_inertly() {
        let xml = wrap(&format!(
            r#"<p:extLst><p:ext uri="{SLIDE_GUIDES_URI}"><p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="7" name="Named" orient="vert" pos="-20" userDrawn="0"><p15:clr><a:schemeClr val="accent1"/></p15:clr><p:extLst><p:ext uri="urn:guide"><v:data xmlns:v="urn:vendor" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></p:ext></p:extLst></p15:guide><p:extLst><p:ext uri="urn:list"><v:list xmlns:v="urn:vendor"/></p:ext></p:extLst></p15:sldGuideLst></p:ext><p:ext uri="{NOTES_GUIDES_URI}"><p15:notesGuideLst xmlns:p15="{P15}"/></p:ext></p:extLst>"#
        ));
        let value = Guides::from_xml(xml.as_bytes(), Limits::DEFAULT).unwrap();
        let guide = &value.slide.as_ref().unwrap().guides[0];
        assert_eq!(guide.name.as_deref(), Some("Named"));
        assert_eq!(guide.color.kind, ColorKind::Scheme);
        let opaque = std::str::from_utf8(guide.extension_xml.as_deref().unwrap()).unwrap();
        assert!(opaque.contains("rIdNeverFetched"));
        assert!(opaque.contains("https://example.invalid/not-opened"));
        assert!(value.notes.as_ref().unwrap().guides.is_empty());
        for strict in [false, true] {
            let written = value.to_xml(strict, Limits::DEFAULT).unwrap();
            let again = Guides::from_xml(wrap(&written).as_bytes(), Limits::DEFAULT).unwrap();
            assert_eq!(again.slide.as_ref().unwrap().guides[0].id, 7);
            assert!(
                std::str::from_utf8(
                    again.slide.as_ref().unwrap().guides[0]
                        .extension_xml
                        .as_deref()
                        .unwrap()
                )
                .unwrap()
                .contains("rIdNeverFetched")
            );
        }
    }

    #[test]
    fn rejects_hostile_extended_guide_grammar() {
        let known = |body: &str| {
            wrap(&format!(
                r#"<p:extLst><p:ext uri="{SLIDE_GUIDES_URI}">{body}</p:ext></p:extLst>"#
            ))
        };
        let cases = [
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="1" orient="diagonal"><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="1"><p15:clr/></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="1"><p15:clr><a:srgbClr val="AABBCC"/><a:schemeClr val="accent1"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p15:guide id="1" r:id="rIdNo"><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="1"><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide><p15:guide id="1"><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(r#"<p15:notesGuideLst xmlns:p15="{P15}"/>"#)),
            format!(r#"<!DOCTYPE x><p:presentation xmlns:p="{P}"/>"#),
        ];
        for xml in cases {
            assert!(
                Guides::from_xml(xml.as_bytes(), Limits::DEFAULT).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_programmatic_limits_and_color_mismatch() {
        let color = Color {
            kind: ColorKind::Srgb,
            xml: format!(r#"<a:srgbClr xmlns:a="{A}" val="AABBCC"/>"#).into_bytes(),
        };
        let guide = Guide {
            id: 1,
            name: None,
            orientation: None,
            position: None,
            user_drawn: None,
            color: color.clone(),
            extension_xml: None,
        };
        let mut value = Guides {
            slide: Some(List {
                guides: vec![guide; Limits::DEFAULT.max_guides() + 1],
                extension_xml: None,
            }),
            notes: None,
        };
        assert!(value.to_xml(false, Limits::DEFAULT).is_err());
        value.slide.as_mut().unwrap().guides.truncate(1);
        value.slide.as_mut().unwrap().guides[0].color.kind = ColorKind::Scheme;
        assert!(value.to_xml(false, Limits::DEFAULT).is_err());
    }

    #[test]
    fn empty_guide_depth_is_checked_in_parse_and_source_scan() {
        let limits = Limits::DEFAULT.with_depth(1).expect("depth limit");
        let exact = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        Guides::from_xml(exact, limits).expect("depth-one empty root");
        rewrite_source(exact, &Guides::default(), limits).expect("depth-one empty source root");

        let one_over = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:holder/></p:presentation>"#;
        assert_xml_depth_limit(
            Guides::from_xml(one_over, limits).expect_err("depth-two empty element"),
            2,
            1,
        );
        assert_xml_depth_limit(
            rewrite_source(one_over, &Guides::default(), limits)
                .expect_err("depth-two source element"),
            2,
            1,
        );
        assert!(matches!(
            Guides::from_xml(
                b"<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                Limits::DEFAULT,
            ),
            Err(Error::Invalid(_))
        ));
    }

    fn assert_xml_depth_limit(error: Error, observed: usize, maximum: usize) {
        match error {
            Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                observed: actual,
                maximum: configured,
            } => {
                assert_eq!(actual, observed);
                assert_eq!(configured, maximum);
            },
            other => panic!("expected XML depth limit, got {other:?}"),
        }
    }
}
