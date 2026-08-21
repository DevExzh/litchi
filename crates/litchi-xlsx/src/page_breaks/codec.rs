//! Bounded `SpreadsheetML` page-break parsing and byte-minimal rewriting.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "the streaming parser keeps wire-state declarations and transitions in scan order"
)]

use std::ops::Range;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};

use super::{Axis, Break, Collection, PageBreaks};
use crate::error::{Error, Result, allocation, invalid};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;

/// Parse the worksheet's direct horizontal and vertical page breaks.
///
/// # Errors
///
/// Returns a typed error for malformed, unsafe, out-of-grid, mismatched-count,
/// or resource-exhausting worksheet XML.
pub fn parse(xml: &[u8]) -> Result<PageBreaks> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("page-break worksheet XML exceeds the size limit"));
    }
    let limits = Limits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..Limits::default()
    };
    let processed = process_markup_compatibility(xml, &Capabilities::default(), &limits)?;
    if processed.xml.len() > MAX_XML_BYTES {
        return Err(invalid("processed page-break XML exceeds the size limit"));
    }
    let selected = if processed.report.alternate_content_count == 0 {
        xml
    } else {
        processed.xml.as_ref()
    };
    parse_selected(selected).map(|parsed| parsed.value)
}

/// Serialize both optional collections as a byte-minimal worksheet fragment.
///
/// # Errors
///
/// Returns an error when the supplied model violates its axis or grid bounds.
pub fn write(value: &PageBreaks) -> Result<Vec<u8>> {
    value.validate()?;
    write_fragment(value, None)
}

/// Replace, insert, or remove direct page-break elements without rebuilding
/// any unrelated worksheet XML.
///
/// # Errors
///
/// Returns an error for invalid source XML, invalid staged values, incompatible
/// MCE projection, or an output resource-limit violation.
pub fn replace(xml: &[u8], value: &PageBreaks) -> Result<Vec<u8>> {
    value.validate()?;
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("page-break worksheet XML exceeds the size limit"));
    }
    let direct = parse_selected(xml)?;
    let limits = Limits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..Limits::default()
    };
    let processed = process_markup_compatibility(xml, &Capabilities::default(), &limits)?;
    if processed.report.alternate_content_count != 0 {
        let selected = parse_selected(processed.xml.as_ref())?;
        if selected.value != direct.value {
            if selected.value == *value {
                return copy_bytes(xml);
            }
            return Err(invalid(
                "page breaks projected through markup compatibility cannot be edited",
            ));
        }
    }

    if direct.value == *value {
        return copy_bytes(xml);
    }
    if direct.lossy_collection {
        return Err(invalid(
            "changed page-break collections contain comments, whitespace, namespace declarations, or qualified attributes",
        ));
    }
    let replacement = write_fragment(value, direct.worksheet_prefix.as_deref())?;
    let insertion = direct
        .first_break_start
        .or(direct.successor_start)
        .or(direct.root_close)
        .ok_or_else(|| invalid("worksheet has no page-break insertion point"))?;
    let mut ranges = [direct.horizontal_span, direct.vertical_span]
        .into_iter()
        .flatten()
        .collect::<Vec<_>>();
    ranges.sort_unstable_by_key(|range| range.start);
    let removed = ranges.iter().try_fold(0usize, |total, range| {
        total
            .checked_add(range.len())
            .ok_or_else(|| invalid("page-break replacement size overflow"))
    })?;
    let capacity = xml
        .len()
        .checked_sub(removed)
        .and_then(|value| value.checked_add(replacement.len()))
        .ok_or_else(|| invalid("page-break replacement size overflow"))?;
    if capacity > MAX_XML_BYTES {
        return Err(invalid("rewritten page-break XML exceeds the size limit"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("page-break worksheet output", source))?;
    let mut cursor = 0usize;
    let mut inserted = false;
    for range in &ranges {
        if !inserted && insertion <= range.start {
            output.extend_from_slice(&xml[cursor..insertion]);
            output.extend_from_slice(&replacement);
            cursor = insertion;
            inserted = true;
        }
        output.extend_from_slice(&xml[cursor..range.start]);
        cursor = range.end;
    }
    if !inserted {
        output.extend_from_slice(&xml[cursor..insertion]);
        output.extend_from_slice(&replacement);
        cursor = insertion;
    }
    output.extend_from_slice(&xml[cursor..]);
    crate::raw::compact::changed(&output, "compact page-break worksheet output")
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Dialect {
    Transitional,
    Strict,
}

#[derive(Debug)]
struct Parsed {
    value: PageBreaks,
    worksheet_prefix: Option<Box<[u8]>>,
    lossy_collection: bool,
    horizontal_span: Option<Range<usize>>,
    vertical_span: Option<Range<usize>>,
    first_break_start: Option<usize>,
    successor_start: Option<usize>,
    root_close: Option<usize>,
}

#[derive(Debug)]
struct OpenCollection {
    axis: Axis,
    start: usize,
    count: Option<usize>,
    manual_count: Option<usize>,
    breaks: Vec<Break>,
}

#[derive(Debug)]
struct OpenBreak {
    depth: usize,
    value: Break,
}

fn parse_selected(xml: &[u8]) -> Result<Parsed> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut events = 0usize;
    let mut value = PageBreaks::new();
    let mut dialect = None;
    let mut worksheet_prefix = None;
    let mut lossy_collection = false;
    let mut horizontal_span = None;
    let mut vertical_span = None;
    let mut first_break_start: Option<usize> = None;
    let mut successor_start: Option<usize> = None;
    let mut root_close = None;
    let mut collection: Option<OpenCollection> = None;
    let mut open_break: Option<OpenBreak> = None;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("page-break XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("page-break XML exceeds the event limit"));
        }
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        reject_unknown_namespace(&namespace)?;
        if let Event::Start(element) | Event::Empty(element) = &event {
            reject_unknown_attribute_prefixes(&reader, element)?;
        }
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("page-break XML contains content after its root"));
                }
                if depth == 0 {
                    if root_seen
                        || dialect_of(&namespace).is_none()
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("page-break parser requires one worksheet root"));
                    }
                    dialect = dialect_of(&namespace);
                    worksheet_prefix = qualified_prefix(element.name().as_ref());
                    root_seen = true;
                } else if let Some(expected) = dialect {
                    reject_dialect_mismatch(&namespace, expected)?;
                }
                if depth == 1 {
                    let local = element.local_name();
                    if is_page_break_collection(local.as_ref())
                        && !same_dialect(&namespace, dialect)
                    {
                        return Err(invalid(
                            "page-break collection namespace does not match worksheet dialect",
                        ));
                    }
                    if same_dialect(&namespace, dialect) {
                        match local.as_ref() {
                            b"rowBreaks" => {
                                if value.horizontal().is_some() || collection.is_some() {
                                    return Err(invalid("duplicate rowBreaks element"));
                                }
                                first_break_start =
                                    Some(first_break_start.map_or(start, |v| v.min(start)));
                                collection = Some(begin_collection(
                                    Axis::Horizontal,
                                    start,
                                    &element,
                                    decoder,
                                    &reader,
                                    &mut lossy_collection,
                                )?);
                            },
                            b"colBreaks" => {
                                if value.vertical().is_some() || collection.is_some() {
                                    return Err(invalid("duplicate colBreaks element"));
                                }
                                first_break_start =
                                    Some(first_break_start.map_or(start, |v| v.min(start)));
                                collection = Some(begin_collection(
                                    Axis::Vertical,
                                    start,
                                    &element,
                                    decoder,
                                    &reader,
                                    &mut lossy_collection,
                                )?);
                            },
                            local if break_successor(local) && successor_start.is_none() => {
                                successor_start = Some(start);
                            },
                            _ => {},
                        }
                    }
                } else if let Some(open) = collection.as_mut() {
                    if depth == 2
                        && same_dialect(&namespace, dialect)
                        && element.local_name().as_ref() == b"brk"
                    {
                        if open_break.is_some() {
                            return Err(invalid("nested page break"));
                        }
                        open_break = Some(OpenBreak {
                            depth: depth + 1,
                            value: parse_break(
                                &element,
                                decoder,
                                open.axis,
                                &reader,
                                &mut lossy_collection,
                            )?,
                        });
                    } else {
                        return Err(invalid("unexpected child in page-break collection"));
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("page-break XML nesting overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("page-break XML nesting is too deep"));
                }
            },
            Event::Empty(element) => {
                if !root_seen || root_closed || depth == 0 {
                    return Err(invalid(
                        "page-break XML element is outside the worksheet root",
                    ));
                }
                if let Some(expected) = dialect {
                    reject_dialect_mismatch(&namespace, expected)?;
                }
                if depth == 1 {
                    let local = element.local_name();
                    if is_page_break_collection(local.as_ref())
                        && !same_dialect(&namespace, dialect)
                    {
                        return Err(invalid(
                            "page-break collection namespace does not match worksheet dialect",
                        ));
                    }
                    let axis = match local.as_ref() {
                        b"rowBreaks" if same_dialect(&namespace, dialect) => Some(Axis::Horizontal),
                        b"colBreaks" if same_dialect(&namespace, dialect) => Some(Axis::Vertical),
                        local if break_successor(local) && successor_start.is_none() => {
                            successor_start = Some(start);
                            None
                        },
                        _ => None,
                    };
                    if let Some(axis) = axis {
                        if collection.is_some()
                            || match axis {
                                Axis::Horizontal => value.horizontal().is_some(),
                                Axis::Vertical => value.vertical().is_some(),
                            }
                        {
                            return Err(invalid("duplicate page-break collection"));
                        }
                        first_break_start = Some(first_break_start.map_or(start, |v| v.min(start)));
                        let open = begin_collection(
                            axis,
                            start,
                            &element,
                            decoder,
                            &reader,
                            &mut lossy_collection,
                        )?;
                        let built = finish_collection(open)?;
                        match axis {
                            Axis::Horizontal => {
                                value.set_horizontal(built)?;
                                horizontal_span = Some(start..end);
                            },
                            Axis::Vertical => {
                                value.set_vertical(built)?;
                                vertical_span = Some(start..end);
                            },
                        }
                    }
                } else if let Some(open) = collection.as_mut() {
                    if depth == 2
                        && same_dialect(&namespace, dialect)
                        && element.local_name().as_ref() == b"brk"
                    {
                        push_break(
                            open,
                            parse_break(
                                &element,
                                decoder,
                                open.axis,
                                &reader,
                                &mut lossy_collection,
                            )?,
                        )?;
                    } else {
                        return Err(invalid("unexpected child in page-break collection"));
                    }
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected page-break XML end element"));
                }
                if let Some(expected) = dialect {
                    reject_dialect_mismatch(&namespace, expected)?;
                }
                if open_break.as_ref().is_some_and(|open| open.depth == depth) {
                    if !same_dialect(&namespace, dialect) || element.local_name().as_ref() != b"brk"
                    {
                        return Err(invalid("invalid page-break closing element"));
                    }
                    let item = open_break
                        .take()
                        .ok_or_else(|| invalid("invalid page-break parser state"))?;
                    push_break(
                        collection
                            .as_mut()
                            .ok_or_else(|| invalid("page break has no collection"))?,
                        item.value,
                    )?;
                } else if depth == 2 && collection.is_some() {
                    let open = collection
                        .take()
                        .ok_or_else(|| invalid("invalid page-break collection state"))?;
                    let expected = match open.axis {
                        Axis::Horizontal => b"rowBreaks".as_slice(),
                        Axis::Vertical => b"colBreaks".as_slice(),
                    };
                    if !same_dialect(&namespace, dialect)
                        || element.local_name().as_ref() != expected
                    {
                        return Err(invalid("invalid page-break collection closing element"));
                    }
                    let axis = open.axis;
                    let span = open.start..end;
                    let built = finish_collection(open)?;
                    match axis {
                        Axis::Horizontal => {
                            value.set_horizontal(built)?;
                            horizontal_span = Some(span);
                        },
                        Axis::Vertical => {
                            value.set_vertical(built)?;
                            vertical_span = Some(span);
                        },
                    }
                } else if depth == 1 {
                    if !same_dialect(&namespace, dialect)
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("invalid worksheet closing element"));
                    }
                    root_close = Some(start);
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::Text(text) => {
                let whitespace = text.as_ref().iter().all(u8::is_ascii_whitespace);
                if collection.is_some() {
                    if whitespace {
                        lossy_collection = true;
                    } else {
                        return Err(invalid("unexpected text in page-break XML"));
                    }
                } else if ((!root_seen || root_closed) || depth == 1) && !whitespace {
                    return Err(invalid("unexpected text in page-break XML"));
                }
            },
            Event::CData(text) if collection.is_some() => {
                if text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    lossy_collection = true;
                } else {
                    return Err(invalid("unexpected CDATA in page-break XML"));
                }
            },
            Event::CData(_) if depth == 1 || !root_seen || root_closed => {
                return Err(invalid("unexpected CDATA in page-break XML"));
            },
            Event::GeneralRef(_) if collection.is_some() => {
                lossy_collection = true;
            },
            Event::GeneralRef(_) if depth == 1 || !root_seen || root_closed => {
                return Err(invalid("unexpected entity text in page-break XML"));
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid page-break XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Comment(_) if collection.is_some() => {
                lossy_collection = true;
            },
            Event::Comment(_) | Event::CData(_) | Event::GeneralRef(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || collection.is_some() || open_break.is_some() {
        return Err(invalid("unterminated page-break worksheet XML"));
    }
    Ok(Parsed {
        value,
        worksheet_prefix,
        lossy_collection,
        horizontal_span,
        vertical_span,
        first_break_start,
        successor_start,
        root_close,
    })
}

fn reject_unknown_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if let ResolveResult::Unknown(prefix) = namespace {
        return Err(invalid(format!(
            "undeclared XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        )));
    }
    Ok(())
}

fn reject_unknown_attribute_prefixes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        if let ResolveResult::Unknown(prefix) = reader.resolver().resolve_attribute(attribute.key).0
        {
            return Err(invalid(format!(
                "undeclared XML attribute prefix '{}'",
                String::from_utf8_lossy(&prefix)
            )));
        }
    }
    Ok(())
}

fn collection_metadata(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    lossy_collection: &mut bool,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            *lossy_collection = true;
            continue;
        }
        match reader.resolver().resolve_attribute(attribute.key).0 {
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "undeclared XML attribute prefix '{}'",
                    String::from_utf8_lossy(&prefix)
                )));
            },
            ResolveResult::Bound(_) => *lossy_collection = true,
            ResolveResult::Unbound => {},
        }
    }
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn dialect_of(namespace: &ResolveResult<'_>) -> Option<Dialect> {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == CORE => Some(Dialect::Transitional),
        ResolveResult::Bound(value) if value.as_ref() == STRICT => Some(Dialect::Strict),
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => None,
    }
}

fn same_dialect(namespace: &ResolveResult<'_>, expected: Option<Dialect>) -> bool {
    dialect_of(namespace) == expected
}

fn reject_dialect_mismatch(namespace: &ResolveResult<'_>, expected: Dialect) -> Result<()> {
    if dialect_of(namespace).is_some_and(|actual| actual != expected) {
        return Err(invalid(
            "worksheet child namespace does not match worksheet root dialect",
        ));
    }
    Ok(())
}

fn is_page_break_collection(local: &[u8]) -> bool {
    matches!(local, b"rowBreaks" | b"colBreaks")
}

fn qualified_prefix(name: &[u8]) -> Option<Box<[u8]>> {
    name.iter()
        .position(|byte| *byte == b':')
        .map(|position| name[..position].to_vec().into_boxed_slice())
}

fn begin_collection(
    axis: Axis,
    start: usize,
    element: &BytesStart<'_>,
    decoder: Decoder,
    reader: &NsReader<&[u8]>,
    lossy_collection: &mut bool,
) -> Result<OpenCollection> {
    collection_metadata(element, reader, lossy_collection)?;
    let mut count = None;
    let mut manual_count = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref().contains(&b':') || attribute.key.as_ref() == b"xmlns" {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        let target = match attribute.key.local_name().as_ref() {
            b"count" => &mut count,
            b"manualBreakCount" => &mut manual_count,
            name => {
                return Err(invalid(format!(
                    "unknown page-break collection attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        if target.is_some() {
            return Err(invalid("duplicate page-break collection attribute"));
        }
        *target = Some(parse_usize(&value, "page-break collection count")?);
    }
    Ok(OpenCollection {
        axis,
        start,
        count,
        manual_count,
        breaks: Vec::new(),
    })
}

fn parse_break(
    element: &BytesStart<'_>,
    decoder: Decoder,
    axis: Axis,
    reader: &NsReader<&[u8]>,
    lossy_collection: &mut bool,
) -> Result<Break> {
    collection_metadata(element, reader, lossy_collection)?;
    let mut id = None;
    let mut minimum = None;
    let mut maximum = None;
    let mut manual = None;
    let mut pivot = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref().contains(&b':') || attribute.key.as_ref() == b"xmlns" {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        match attribute.key.local_name().as_ref() {
            b"id" => set_once(&mut id, parse_u32(&value, "page-break id")?)?,
            b"min" => set_once(&mut minimum, parse_u32(&value, "page-break minimum")?)?,
            b"max" => set_once(&mut maximum, parse_u32(&value, "page-break maximum")?)?,
            b"man" => set_once(&mut manual, parse_bool(&value)?)?,
            b"pt" => set_once(&mut pivot, parse_bool(&value)?)?,
            name => {
                return Err(invalid(format!(
                    "unknown page-break attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        }
    }
    let value = Break::new(id.unwrap_or(0), minimum.unwrap_or(0), maximum.unwrap_or(0))?
        .with_manual(manual.unwrap_or(false))
        .with_pivot(pivot.unwrap_or(false));
    value.validate(axis)?;
    Ok(value)
}

fn set_once<T>(target: &mut Option<T>, value: T) -> Result<()> {
    if target.is_some() {
        return Err(invalid("duplicate page-break attribute"));
    }
    *target = Some(value);
    Ok(())
}

fn push_break(collection: &mut OpenCollection, value: Break) -> Result<()> {
    let maximum = match collection.axis {
        Axis::Horizontal => super::MAX_HORIZONTAL_BREAKS,
        Axis::Vertical => super::MAX_VERTICAL_BREAKS,
    };
    if collection.breaks.len() >= maximum {
        return Err(invalid("page-break collection exceeds the Office limit"));
    }
    collection
        .breaks
        .try_reserve(1)
        .map_err(|source| allocation("worksheet page breaks", source))?;
    collection.breaks.push(value);
    Ok(())
}

fn finish_collection(open: OpenCollection) -> Result<Collection> {
    let actual = open.breaks.len();
    if open.count.unwrap_or(0) != actual {
        return Err(invalid("page-break count does not match its children"));
    }
    let manual = open.breaks.iter().filter(|value| value.is_manual()).count();
    if open.manual_count.unwrap_or(0) != manual {
        return Err(invalid(
            "manualBreakCount does not match manual page breaks",
        ));
    }
    match open.axis {
        Axis::Horizontal => Collection::horizontal(open.breaks),
        Axis::Vertical => Collection::vertical(open.breaks),
    }
}

fn write_fragment(value: &PageBreaks, prefix: Option<&[u8]>) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    if let Some(collection) = value.horizontal() {
        write_collection(&mut output, "rowBreaks", collection, prefix)?;
    }
    if let Some(collection) = value.vertical() {
        write_collection(&mut output, "colBreaks", collection, prefix)?;
    }
    Ok(output)
}

fn write_collection(
    output: &mut Vec<u8>,
    name: &str,
    collection: &Collection,
    prefix: Option<&[u8]>,
) -> Result<()> {
    let expected = match name {
        "rowBreaks" => Axis::Horizontal,
        "colBreaks" => Axis::Vertical,
        _ => return Err(invalid("invalid page-break collection name")),
    };
    if collection.axis() != expected {
        return Err(invalid("page-break collection axis mismatch"));
    }
    output.extend_from_slice(b"<");
    if let Some(prefix) = prefix {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(name.as_bytes());
    if !collection.is_empty() {
        let count = u32::try_from(collection.len()).map_err(|error| {
            invalid(format!(
                "page-break count does not fit unsignedInt: {error}"
            ))
        })?;
        push_number_attribute(output, "count", count);
    }
    let manual = collection.manual_break_count();
    if manual != 0 {
        let manual = u32::try_from(manual).map_err(|error| {
            invalid(format!(
                "manual page-break count does not fit unsignedInt: {error}"
            ))
        })?;
        push_number_attribute(output, "manualBreakCount", manual);
    }
    if collection.is_empty() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.extend_from_slice(b">");
    for value in collection.breaks() {
        if let Some(prefix) = prefix {
            output.extend_from_slice(b"<");
            output.extend_from_slice(prefix);
            output.extend_from_slice(b":brk");
        } else {
            output.extend_from_slice(b"<brk");
        }
        if value.id() != 0 {
            push_number_attribute(output, "id", value.id());
        }
        if value.minimum() != 0 {
            push_number_attribute(output, "min", value.minimum());
        }
        if value.maximum() != 0 {
            push_number_attribute(output, "max", value.maximum());
        }
        if value.is_manual() {
            output.extend_from_slice(b" man=\"1\"");
        }
        if value.is_pivot() {
            output.extend_from_slice(b" pt=\"1\"");
        }
        output.extend_from_slice(b"/>");
    }
    output.extend_from_slice(b"</");
    if let Some(prefix) = prefix {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
    Ok(())
}

fn push_number_attribute(output: &mut Vec<u8>, name: &str, value: u32) {
    output.extend_from_slice(b" ");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    let mut buffer = itoa::Buffer::new();
    output.extend_from_slice(buffer.format(value).as_bytes());
    output.extend_from_slice(b"\"");
}

fn break_successor(local: &[u8]) -> bool {
    matches!(
        local,
        b"customProperties"
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

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "0" | "false" => Ok(false),
        "1" | "true" => Ok(true),
        _ => Err(invalid("invalid page-break boolean")),
    }
}

fn parse_u32(value: &str, what: &'static str) -> Result<u32> {
    value
        .parse()
        .map_err(|error| invalid(format!("invalid {what}: {error}")))
}

fn parse_usize(value: &str, what: &'static str) -> Result<usize> {
    let value = parse_u32(value, what)?;
    usize::try_from(value).map_err(|error| invalid(format!("invalid {what}: {error}")))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|error| {
        invalid(format!(
            "page-break XML position does not fit usize: {error}"
        ))
    })
}

fn copy_bytes(input: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(input.len())
        .map_err(|source| allocation("page-break worksheet output", source))?;
    output.extend_from_slice(input);
    Ok(output)
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid worksheet page-break XML: {error}"
    )))
}
