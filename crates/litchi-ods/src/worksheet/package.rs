//! Package-boundary replacement of direct spreadsheet tables.
//!
//! The package layer edits only the table spans owned by
//! `office:spreadsheet`.  Unmodelled spreadsheet children make a semantic
//! replacement unsafe, so the operation rejects them instead of silently
//! discarding producer data.

use super::{Sheet, codec, validation};
use litchi_core::{Error, Result};
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const MAX_SPANS: usize = 1_048_576;
const OFFICE_NAMESPACE: &str = codec::OFFICE_NAMESPACE;
const TABLE_NAMESPACE: &str = codec::TABLE_NAMESPACE;

#[derive(Clone, Debug)]
struct Span {
    namespace: Option<String>,
    local: String,
    qname: String,
    start: usize,
    tag_end: usize,
    close_start: usize,
    end: usize,
    parent: Option<usize>,
    empty: bool,
}

/// Replace the direct table children of `office:spreadsheet` in one bounded
/// XML pass.  The returned content is structurally and semantically rechecked
/// before publication by the caller.
pub(crate) fn replace_tables(xml: &str, sheets: &[Sheet]) -> Result<String> {
    validation::validate_content_xml_size(xml)?;
    validation::validate_sheets(sheets)?;
    let spans = scan(xml)?;
    let spreadsheet = one_spreadsheet(&spans)?;
    let mut tables = direct_children(&spans, spreadsheet, TABLE_NAMESPACE, "table");
    tables.sort_unstable_by_key(|index| spans[*index].start);

    for table in &tables {
        for child in direct_children_any(&spans, *table) {
            if !is_element(&spans[child], TABLE_NAMESPACE, "table-row") {
                return Err(Error::InvalidFormat(
                    "ODS worksheet edit cannot replace a table containing unmodeled direct children"
                        .to_string(),
                ));
            }
        }
    }

    let mut rendered = Vec::with_capacity(sheets.len());
    for sheet in sheets {
        rendered.push(codec::write_sheet(sheet)?);
    }

    let mut edits = Vec::<(usize, usize, String)>::with_capacity(
        tables.len().max(rendered.len()).saturating_add(1),
    );
    for (position, table) in tables.iter().enumerate() {
        let replacement = rendered.get(position).cloned().unwrap_or_default();
        edits.push((spans[*table].start, spans[*table].end, replacement));
    }

    if rendered.len() > tables.len() {
        let extras = rendered[tables.len()..].join("");
        if spans[spreadsheet].empty {
            let opening = xml
                .get(spans[spreadsheet].start..spans[spreadsheet].tag_end)
                .ok_or_else(|| invalid("invalid spreadsheet start span"))?;
            let opening = opening
                .trim_end()
                .strip_suffix("/>")
                .ok_or_else(|| invalid("empty spreadsheet has no self-closing tag"))?;
            let mut replacement = String::with_capacity(opening.len() + extras.len() + 32);
            replacement.push_str(opening);
            replacement.push('>');
            replacement.push_str(&extras);
            replacement.push_str("</");
            replacement.push_str(&spans[spreadsheet].qname);
            replacement.push('>');
            edits.push((
                spans[spreadsheet].start,
                spans[spreadsheet].tag_end,
                replacement,
            ));
        } else {
            let insertion = direct_children_any(&spans, spreadsheet)
                .into_iter()
                .filter(|child| !is_element(&spans[*child], TABLE_NAMESPACE, "table"))
                .map(|child| spans[child].start)
                .min()
                .unwrap_or(spans[spreadsheet].close_start);
            edits.push((insertion, insertion, extras));
        }
    }

    let updated = apply_edits(xml, edits)?;
    crate::authoring::validate_content_xml(&updated)?;
    codec::parse(&updated)?;
    Ok(updated)
}

fn scan(xml: &str) -> Result<Vec<Span>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut spans = Vec::new();
    let mut open = Vec::<usize>::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODS XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let element = element.into_owned();
                let parent = open.last().copied();
                let namespace = resolve_namespace(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element local name")?;
                let qname = decode(element.name().as_ref(), "element qualified name")?;
                let index = push_span(
                    xml, &mut spans, &reader, namespace, local, qname, parent, false,
                )?;
                open.push(index);
            },
            Event::Empty(element) => {
                let element = element.into_owned();
                let parent = open.last().copied();
                let namespace = resolve_namespace(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element local name")?;
                let qname = decode(element.name().as_ref(), "element qualified name")?;
                push_span(
                    xml, &mut spans, &reader, namespace, local, qname, parent, true,
                )?;
            },
            Event::End(_) => {
                let index = open.pop().ok_or_else(|| {
                    Error::InvalidFormat("ODS XML element stack underflow".to_string())
                })?;
                let end = position(&reader)?;
                let close_start = xml.as_bytes()[..end]
                    .windows(2)
                    .rposition(|bytes| bytes == b"</")
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODS XML closing tag start is missing".to_string())
                    })?;
                spans[index].close_start = close_start;
                spans[index].end = end;
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    if !open.is_empty() {
        return Err(Error::InvalidFormat(
            "ODS XML contains unclosed elements".to_string(),
        ));
    }
    Ok(spans)
}

fn push_span(
    xml: &str,
    spans: &mut Vec<Span>,
    reader: &NsReader<&[u8]>,
    namespace: Option<String>,
    local: String,
    qname: String,
    parent: Option<usize>,
    empty: bool,
) -> Result<usize> {
    if spans.len() >= MAX_SPANS {
        return Err(Error::InvalidFormat(
            "ODS content contains too many XML elements".to_string(),
        ));
    }
    let tag_end = position(reader)?;
    let index = spans.len();
    spans.push(Span {
        namespace,
        local,
        qname,
        start: tag_start(xml, tag_end)?,
        tag_end,
        close_start: tag_end,
        end: tag_end,
        parent,
        empty,
    });
    Ok(index)
}

fn tag_start(xml: &str, tag_end: usize) -> Result<usize> {
    xml.as_bytes()[..tag_end]
        .iter()
        .rposition(|&byte| byte == b'<')
        .ok_or_else(|| Error::InvalidFormat("ODS XML element start is missing".to_string()))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| Error::InvalidFormat("ODS XML position overflows usize".to_string()))
}

fn resolve_namespace(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => Ok(Some(decode(uri, "element namespace")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound ODS element prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn decode(bytes: &[u8], label: &str) -> Result<String> {
    String::from_utf8(bytes.to_vec())
        .map_err(|_| Error::InvalidFormat(format!("ODS {label} is not valid UTF-8")))
}

fn one_spreadsheet(spans: &[Span]) -> Result<usize> {
    let mut matches = spans
        .iter()
        .enumerate()
        .filter(|(_, span)| is_element(span, OFFICE_NAMESPACE, "spreadsheet"));
    let spreadsheet = matches
        .next()
        .map(|(index, _)| index)
        .ok_or_else(|| invalid("ODS content has no office:spreadsheet"))?;
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(
            "ODS content has more than one office:spreadsheet".to_string(),
        ));
    }
    Ok(spreadsheet)
}

fn is_element(span: &Span, namespace: &str, local: &str) -> bool {
    span.namespace.as_deref() == Some(namespace) && span.local == local
}

fn direct_children(spans: &[Span], parent: usize, namespace: &str, local: &str) -> Vec<usize> {
    spans
        .iter()
        .enumerate()
        .filter_map(|(index, span)| {
            (span.parent == Some(parent) && is_element(span, namespace, local)).then_some(index)
        })
        .collect()
}

fn direct_children_any(spans: &[Span], parent: usize) -> Vec<usize> {
    spans
        .iter()
        .enumerate()
        .filter_map(|(index, span)| (span.parent == Some(parent)).then_some(index))
        .collect()
}

fn apply_edits(xml: &str, mut edits: Vec<(usize, usize, String)>) -> Result<String> {
    edits.sort_unstable_by_key(|(start, _, _)| *start);
    let mut output = String::with_capacity(xml.len());
    let mut cursor = 0usize;
    for (start, end, replacement) in edits {
        if start < cursor || end < start || end > xml.len() {
            return Err(Error::InvalidFormat(
                "overlapping or out-of-bounds ODS worksheet edit".to_string(),
            ));
        }
        output.push_str(&xml[cursor..start]);
        output.push_str(&replacement);
        cursor = end;
    }
    output.push_str(&xml[cursor..]);
    Ok(output)
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_string())
}
