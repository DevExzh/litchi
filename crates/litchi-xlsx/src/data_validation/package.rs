//! Worksheet-level data-validation replacement and byte-preserving edits.

use super::codec::{
    BoundedXml, append_bounded_bytes, exact, invalid, optional_attr,
    parse_data_validation_collections, reserve_vec, spreadsheet,
    validate_data_validation_collections, write_data_validation_core,
    write_data_validation_extensions, xml_error,
};
use super::model::{Collection, Conformance, Source};
use super::{
    CORE, EXTENSION_URI, MAX_CAPTURED_COLLECTIONS, MAX_DEPTH, MAX_EVENTS, MAX_NODES, MAX_XML_BYTES,
    STRICT, X14,
};
use crate::error::Result;
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::ops::Range as ByteRange;

#[derive(Debug)]
struct XmlScan {
    conformance: Conformance,
    worksheet_close: usize,
    core_insert: usize,
    core_ranges: Vec<ByteRange<usize>>,
    x14_ranges: Vec<ByteRange<usize>>,
    matching_ext_close: Option<usize>,
    ext_lst_close: Option<usize>,
}

/// Replace data-validation XML while preserving every unrelated worksheet byte.
pub fn replace_data_validation_collections(
    worksheet_xml: &[u8],
    values: &[Collection],
) -> Result<Vec<u8>> {
    let parsed = parse_data_validation_collections(worksheet_xml)?;
    validate_data_validation_collections(&parsed)?;
    validate_data_validation_collections(values)?;
    let scan = scan_data_validation_xml(worksheet_xml)?;
    let parsed_core = parsed.iter().any(|value| value.source == Source::Core);
    let parsed_x14 = parsed
        .iter()
        .any(|value| value.source == Source::Office2010);
    if parsed_core == scan.core_ranges.is_empty() || parsed_x14 == scan.x14_ranges.is_empty() {
        return Err(invalid(
            "data validations selected through MCE cannot be mutated byte-exactly",
        ));
    }
    let core = write_data_validation_core(values, scan.conformance)?;
    let extensions = write_data_validation_extensions(values, scan.conformance)?;
    let edit_count = scan
        .core_ranges
        .len()
        .checked_add(scan.x14_ranges.len())
        .ok_or_else(|| invalid("data-validation edit count overflow"))?;
    let mut edits = Vec::new();
    reserve_vec(&mut edits, edit_count, "data-validation edits")?;
    edits.extend(
        scan.core_ranges
            .iter()
            .chain(scan.x14_ranges.iter())
            .cloned()
            .map(|range| (range, Vec::new())),
    );
    if !core.is_empty() {
        if let Some(range) = scan.core_ranges.first() {
            let Some(edit) = edits.iter_mut().find(|(candidate, _)| candidate == range) else {
                return Err(invalid("missing core data-validation edit"));
            };
            edit.1 = core.into_bytes();
        } else {
            reserve_vec(&mut edits, 1, "data-validation edits")?;
            edits.push((scan.core_insert..scan.core_insert, core.into_bytes()));
        }
    }
    if !extensions.is_empty() {
        let inner = data_validation_extension_inner(&extensions)?;
        if let Some(range) = scan.x14_ranges.first() {
            let Some(edit) = edits.iter_mut().find(|(candidate, _)| candidate == range) else {
                return Err(invalid("missing Office 2010 data-validation edit"));
            };
            edit.1 = inner.into_bytes();
        } else if let Some(position) = scan.matching_ext_close {
            reserve_vec(&mut edits, 1, "data-validation edits")?;
            edits.push((position..position, inner.into_bytes()));
        } else if let Some(position) = scan.ext_lst_close {
            reserve_vec(&mut edits, 1, "data-validation edits")?;
            edits.push((
                position..position,
                data_validation_extension_wrapper(&inner, scan.conformance)?.into_bytes(),
            ));
        } else {
            reserve_vec(&mut edits, 1, "data-validation edits")?;
            edits.push((
                scan.worksheet_close..scan.worksheet_close,
                extensions.into_bytes(),
            ));
        }
    }
    apply_data_validation_edits(worksheet_xml, edits)
}

fn data_validation_extension_inner(fragment: &str) -> Result<String> {
    let start = fragment
        .find("<x14:dataValidations")
        .ok_or_else(|| invalid("invalid generated data-validation extension"))?;
    let end = fragment
        .rfind("</x14:dataValidations>")
        .ok_or_else(|| invalid("invalid generated data-validation extension"))?
        + "</x14:dataValidations>".len();
    Ok(fragment[start..end].to_string())
}

fn data_validation_extension_wrapper(inner: &str, conformance: Conformance) -> Result<String> {
    let mut wrapper = BoundedXml::new();
    wrapper.write_arguments(format_args!(
        "<ext xmlns=\"{}\" uri=\"{}\">{inner}</ext>",
        conformance.namespace(),
        EXTENSION_URI
    ))?;
    Ok(wrapper.finish())
}

fn apply_data_validation_edits(
    xml: &[u8],
    mut edits: Vec<(ByteRange<usize>, Vec<u8>)>,
) -> Result<Vec<u8>> {
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    let mut output = Vec::new();
    reserve_vec(&mut output, xml.len(), "data-validation XML output")?;
    let mut cursor = 0usize;
    for (range, replacement) in edits {
        if range.start < cursor || range.end < range.start || range.end > xml.len() {
            return Err(invalid("overlapping data-validation XML edits"));
        }
        append_bounded_bytes(&mut output, &xml[cursor..range.start])?;
        append_bounded_bytes(&mut output, &replacement)?;
        cursor = range.end;
    }
    append_bounded_bytes(&mut output, &xml[cursor..])?;
    let reparsed = parse_data_validation_collections(&output)?;
    validate_data_validation_collections(&reparsed)?;
    Ok(output)
}

fn push_scan_range(ranges: &mut Vec<ByteRange<usize>>, range: ByteRange<usize>) -> Result<()> {
    if ranges.len() >= MAX_CAPTURED_COLLECTIONS {
        return Err(invalid("too many physical data-validation collections"));
    }
    reserve_vec(ranges, 1, "data-validation scan ranges")?;
    ranges.push(range);
    Ok(())
}

fn scan_data_validation_xml(xml: &[u8]) -> Result<XmlScan> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("data-validation worksheet XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut previous = 0usize;
    let mut conformance = None;
    let mut worksheet_close = None;
    let mut core_insert = None;
    let mut core_start = None;
    let mut core_ranges = Vec::new();
    let mut x14_start = None;
    let mut x14_ranges = Vec::new();
    let mut matching_ext_depth = None;
    let mut matching_ext_close = None;
    let mut ext_lst_depth = None;
    let mut ext_lst_close = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut events = 0usize;
    let mut nodes = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("data-validation worksheet event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("data-validation worksheet exceeds event limit"));
        }
        let start = previous;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_source| invalid("data-validation XML offset overflow"))?;
        previous = end;
        let decoder = reader.decoder();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Eof) {
            break;
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("data-validation worksheet node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("data-validation worksheet exceeds node limit"));
            }
        }
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if depth == 0 && root_seen {
                    return Err(invalid("worksheet XML contains multiple roots"));
                }
                if depth >= MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                let local = element.local_name();
                if depth == 1 {
                    conformance = match namespace {
                        ResolveResult::Bound(value) if value.as_ref() == CORE => {
                            Some(Conformance::Transitional)
                        },
                        ResolveResult::Bound(value) if value.as_ref() == STRICT => {
                            Some(Conformance::Strict)
                        },
                        ResolveResult::Unbound
                        | ResolveResult::Bound(_)
                        | ResolveResult::Unknown(_) => None,
                    };
                    if conformance.is_none() || local.as_ref() != b"worksheet" {
                        return Err(invalid("invalid worksheet namespace"));
                    }
                    root_seen = true;
                } else if depth == 2 {
                    if local.as_ref() == b"dataValidations" && !spreadsheet(&namespace) {
                        return Err(invalid("spoofed dataValidations element namespace"));
                    }
                    if spreadsheet(&namespace) {
                        if local.as_ref() == b"dataValidations" {
                            if core_start.is_some() {
                                return Err(invalid("duplicate core dataValidations element"));
                            }
                            core_start = Some((depth, start));
                        } else if core_insert.is_none() && validation_schema_after(local.as_ref()) {
                            core_insert = Some(start);
                        }
                        if local.as_ref() == b"extLst" {
                            ext_lst_depth = Some(depth);
                        }
                    }
                }
                if spreadsheet(&namespace)
                    && local.as_ref() == b"ext"
                    && optional_attr(&element, b"uri", decoder)?.as_deref() == Some(EXTENSION_URI)
                {
                    if matching_ext_depth.is_some() {
                        return Err(invalid("nested data-validation extension"));
                    }
                    matching_ext_depth = Some(depth);
                }
                if local.as_ref() == b"dataValidations" && matching_ext_depth.is_some() {
                    if !exact(&namespace, X14) {
                        return Err(invalid("spoofed x14 dataValidations element namespace"));
                    }
                    if x14_start.is_some() {
                        return Err(invalid("duplicate Office 2010 dataValidations element"));
                    }
                    x14_start = Some((depth, start));
                }
            },
            Event::Empty(element) => {
                if root_closed || depth == 0 {
                    return Err(invalid("worksheet XML contains an element outside root"));
                }
                let local = element.local_name();
                let element_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if element_depth > MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                if element_depth == 2 {
                    if local.as_ref() == b"dataValidations" && !spreadsheet(&namespace) {
                        return Err(invalid("spoofed dataValidations element namespace"));
                    }
                    if spreadsheet(&namespace) && local.as_ref() == b"dataValidations" {
                        if core_start.is_some() {
                            return Err(invalid("duplicate core dataValidations element"));
                        }
                        push_scan_range(&mut core_ranges, start..end)?;
                    } else if spreadsheet(&namespace)
                        && core_insert.is_none()
                        && validation_schema_after(local.as_ref())
                    {
                        core_insert = Some(start);
                    }
                }
                if local.as_ref() == b"dataValidations" && matching_ext_depth.is_some() {
                    if !exact(&namespace, X14) {
                        return Err(invalid("spoofed x14 dataValidations element namespace"));
                    }
                    push_scan_range(&mut x14_ranges, start..end)?;
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected worksheet closing element"));
                }
                if core_start.is_some_and(|(element_depth, _)| element_depth == depth) {
                    let Some((_, range_start)) = core_start.take() else {
                        return Err(invalid("missing core data-validation scan state"));
                    };
                    push_scan_range(&mut core_ranges, range_start..end)?;
                }
                if x14_start.is_some_and(|(element_depth, _)| element_depth == depth) {
                    let Some((_, range_start)) = x14_start.take() else {
                        return Err(invalid("missing Office 2010 data-validation scan state"));
                    };
                    push_scan_range(&mut x14_ranges, range_start..end)?;
                }
                if matching_ext_depth == Some(depth) {
                    matching_ext_close = Some(start);
                    matching_ext_depth = None;
                }
                if ext_lst_depth == Some(depth) {
                    ext_lst_close = Some(start);
                    ext_lst_depth = None;
                }
                if depth == 1 && element.local_name().as_ref() == b"worksheet" {
                    if !spreadsheet(&namespace) {
                        return Err(invalid("invalid worksheet closing namespace"));
                    }
                    root_closed = true;
                    worksheet_close = Some(start);
                } else if depth == 1 {
                    return Err(invalid("invalid worksheet closing element"));
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("worksheet XML nesting underflow"))?;
            },
            Event::Text(value) => {
                if (!root_seen || root_closed)
                    && !value.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("worksheet XML text is outside root"));
                }
                if depth == 1 && !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("worksheet cannot contain direct text"));
                }
            },
            Event::CData(_) => {
                return Err(invalid("worksheet XML contains unexpected CDATA"));
            },
            Event::GeneralRef(reference) => {
                decode_xml_reference(&reference)?;
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid worksheet XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || core_start.is_some() || x14_start.is_some() {
        return Err(invalid("incomplete worksheet data-validation XML"));
    }
    let worksheet_close = worksheet_close.ok_or_else(|| invalid("worksheet is not closed"))?;
    Ok(XmlScan {
        conformance: conformance.ok_or_else(|| invalid("invalid worksheet namespace"))?,
        worksheet_close,
        core_insert: core_insert.unwrap_or(worksheet_close),
        core_ranges,
        x14_ranges,
        matching_ext_close,
        ext_lst_close,
    })
}

fn validation_schema_after(local: &[u8]) -> bool {
    matches!(
        local,
        b"hyperlinks"
            | b"printOptions"
            | b"pageMargins"
            | b"pageSetup"
            | b"headerFooter"
            | b"rowBreaks"
            | b"colBreaks"
            | b"customProperties"
            | b"cellWatches"
            | b"ignoredErrors"
            | b"smartTags"
            | b"drawing"
            | b"legacyDrawing"
            | b"legacyDrawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    )
}
