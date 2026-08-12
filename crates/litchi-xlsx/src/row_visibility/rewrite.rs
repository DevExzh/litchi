//! Lossless direct `hidden` attribute scanning and rewriting.

use std::collections::BTreeMap;
use std::ops::Range;

use litchi_sheet::Row;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Result, allocation, invalid};

const TRANSITIONAL_SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";

#[derive(Debug)]
struct RowTag {
    row: Row,
    tag: Range<usize>,
    hidden: Option<Range<usize>>,
    canonical_hidden: bool,
    insert_at: usize,
}

pub(super) fn scan(xml: &[u8]) -> Result<Vec<(Row, bool)>> {
    let tags = row_tags(xml)?;
    let mut rows = Vec::new();
    rows.try_reserve_exact(tags.len())
        .map_err(|source| allocation("row-visibility snapshot", source))?;
    for tag in tags {
        rows.push((tag.row, hidden_value(xml, &tag)?));
    }
    Ok(rows)
}

pub(super) fn rewrite(xml: &[u8], actions: &BTreeMap<Row, bool>) -> Result<(Vec<u8>, usize)> {
    if actions.is_empty() {
        return Ok((xml.to_vec(), 0));
    }
    let tags = row_tags(xml)?;
    let mut by_row = BTreeMap::new();
    for tag in tags {
        by_row.insert(tag.row, tag);
    }
    let mut changed = 0usize;
    let mut extra = 0usize;
    for (row, hidden) in actions {
        let tag = by_row
            .get(row)
            .ok_or_else(|| invalid(format!("row-visibility row '{}' disappeared", row)))?;
        let changes = if *hidden {
            !tag.canonical_hidden
        } else {
            tag.hidden.is_some()
        };
        if changes {
            changed += 1;
            extra = extra.saturating_add(11);
        }
    }
    if changed == 0 {
        return Ok((xml.to_vec(), 0));
    }
    let capacity = xml
        .len()
        .checked_add(extra)
        .ok_or_else(|| invalid("row-visibility output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("row-visibility worksheet output", source))?;
    let mut cursor = 0usize;
    for (row, hidden) in actions {
        let tag = by_row
            .get(row)
            .ok_or_else(|| invalid(format!("row-visibility row '{}' disappeared", row)))?;
        if (*hidden && tag.canonical_hidden) || (!*hidden && tag.hidden.is_none()) {
            continue;
        }
        output.extend_from_slice(&xml[cursor..tag.tag.start]);
        write_tag(&mut output, xml, tag, *hidden);
        cursor = tag.tag.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok((output, changed))
}

fn write_tag(output: &mut Vec<u8>, xml: &[u8], tag: &RowTag, hidden: bool) {
    let Some(attribute) = tag.hidden.as_ref() else {
        output.extend_from_slice(&xml[tag.tag.start..tag.insert_at]);
        if hidden {
            output.extend_from_slice(b" hidden=\"1\"");
        }
        output.extend_from_slice(&xml[tag.insert_at..tag.tag.end]);
        return;
    };
    output.extend_from_slice(&xml[tag.tag.start..attribute.start]);
    output.extend_from_slice(&xml[attribute.end..tag.insert_at]);
    if hidden {
        output.extend_from_slice(b" hidden=\"1\"");
    }
    output.extend_from_slice(&xml[tag.insert_at..tag.tag.end]);
}

fn row_tags(xml: &[u8]) -> Result<Vec<RowTag>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut tags = Vec::new();
    let mut previous_row = 0u32;
    let mut inside_sheet_data = false;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_source| invalid("row-visibility XML position does not fit usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("row-visibility XML scan failed: {error}")))?
            .into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_source| invalid("row-visibility XML position does not fit usize"))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_sheet_data(&namespace, &element) => {
                if inside_sheet_data {
                    return Err(invalid("row-visibility worksheet has nested sheetData"));
                }
                inside_sheet_data = true;
            },
            Event::Empty(element) if is_sheet_data(&namespace, &element) => {},
            Event::End(element)
                if is_spreadsheetml_local(
                    &namespace,
                    element.local_name().as_ref(),
                    b"sheetData",
                ) =>
            {
                if !inside_sheet_data {
                    return Err(invalid(
                        "row-visibility worksheet has an unmatched sheetData close",
                    ));
                }
                inside_sheet_data = false;
            },
            Event::Start(element) | Event::Empty(element)
                if inside_sheet_data && is_row(&namespace, &element) =>
            {
                let row = parse_row(&element, reader.decoder(), previous_row)?;
                if row <= previous_row {
                    return Err(invalid(
                        "row-visibility edits require strictly increasing row owners",
                    ));
                }
                previous_row = row;
                let checked = Row::new(row - 1)?;
                let (hidden, canonical_hidden, insert_at) = lexical_tag(xml, start..end)?;
                tags.try_reserve(1)
                    .map_err(|source| allocation("row-visibility row index", source))?;
                tags.push(RowTag {
                    row: checked,
                    tag: start..end,
                    hidden,
                    canonical_hidden,
                    insert_at,
                });
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "row-visibility edits refuse XML document type declarations",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(tags)
}

fn is_row(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    is_spreadsheetml_local(namespace, element.name().local_name().as_ref(), b"row")
}

fn is_sheet_data(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    is_spreadsheetml_local(
        namespace,
        element.name().local_name().as_ref(),
        b"sheetData",
    )
}

fn is_spreadsheetml_local(namespace: &ResolveResult<'_>, actual: &[u8], expected: &[u8]) -> bool {
    actual == expected
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TRANSITIONAL_SML || *value == STRICT_SML)
}

fn parse_row(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    previous: u32,
) -> Result<u32> {
    let mut explicit = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid row-visibility attribute: {error}")))?;
        if attribute.key.as_ref() == b"r" {
            if explicit.is_some() {
                return Err(invalid("duplicate row coordinate attribute"));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| invalid(format!("invalid row coordinate: {error}")))?;
            explicit = Some(crate::raw::worksheet::parse_one_based_row(&value)?);
        }
    }
    previous
        .checked_add(1)
        .filter(|value| *value <= litchi_sheet::ROWS)
        .map_or_else(
            || {
                Err(invalid(
                    "inferred row-visibility coordinate exceeds the grid",
                ))
            },
            |inferred| Ok(explicit.unwrap_or(inferred)),
        )
}

fn lexical_tag(xml: &[u8], tag: Range<usize>) -> Result<(Option<Range<usize>>, bool, usize)> {
    let bytes = xml
        .get(tag.clone())
        .ok_or_else(|| invalid("row-visibility tag span exceeds worksheet XML"))?;
    if bytes.first() != Some(&b'<') || bytes.last() != Some(&b'>') {
        return Err(invalid("row-visibility row tag has invalid boundaries"));
    }
    let mut close = bytes.len() - 1;
    if close > 0 && bytes[close - 1] == b'/' {
        close -= 1;
    }
    let insert_at = tag.start + close;
    let mut cursor = 1usize;
    while cursor < close && !bytes[cursor].is_ascii_whitespace() && bytes[cursor] != b'/' {
        cursor += 1;
    }
    let mut hidden = None;
    let mut canonical = false;
    while cursor < close {
        let leading = cursor;
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor == close {
            break;
        }
        let name_start = cursor;
        while cursor < close
            && !bytes[cursor].is_ascii_whitespace()
            && !matches!(bytes[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor == close || bytes[cursor] != b'=' {
            return Err(invalid("row-visibility row attribute has no value"));
        }
        cursor += 1;
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'\"'))
            .ok_or_else(|| invalid("row-visibility row attribute is not quoted"))?;
        cursor += 1;
        while cursor < close && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor == close {
            return Err(invalid("row-visibility row attribute is not closed"));
        }
        cursor += 1;
        if &bytes[name_start..name_end] == b"hidden" {
            if hidden.is_some() {
                return Err(invalid("duplicate row hidden attribute"));
            }
            hidden = Some(tag.start + leading..tag.start + cursor);
            canonical = &bytes[leading..cursor] == b" hidden=\"1\"";
        }
    }
    Ok((hidden, canonical, insert_at))
}

fn hidden_value(xml: &[u8], tag: &RowTag) -> Result<bool> {
    let Some(range) = tag.hidden.as_ref() else {
        return Ok(false);
    };
    let attribute = xml
        .get(range.clone())
        .ok_or_else(|| invalid("row hidden attribute span exceeds worksheet XML"))?;
    let equals = attribute
        .iter()
        .position(|byte| *byte == b'=')
        .ok_or_else(|| invalid("row hidden attribute has no value"))?;
    let value = std::str::from_utf8(&attribute[equals + 1..])
        .map_err(|error| invalid(format!("row hidden attribute is not UTF-8: {error}")))?
        .trim();
    let bytes = value.as_bytes();
    if bytes.len() < 2 || !matches!(bytes[0], b'\'' | b'\"') || bytes.last() != Some(&bytes[0]) {
        return Err(invalid("row hidden attribute is not quoted"));
    }
    let value = &value[1..value.len() - 1];
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid row hidden value '{value}'"))),
    }
}

#[cfg(test)]
mod tests {
    use super::scan;

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[test]
    fn duplicate_row_attributes_fail_closed() {
        for attributes in [r#"r="1" r="1""#, r#"r="1" hidden="1" hidden="0""#] {
            let xml = format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row {attributes}/></sheetData></worksheet>"#
            );
            assert!(scan(xml.as_bytes()).is_err());
        }
    }
}
