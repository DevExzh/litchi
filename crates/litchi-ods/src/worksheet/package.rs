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

/// Patch only changed physical row runs in flat spreadsheet XML.
///
/// Direct table children outside the changed rows remain byte-exact. A row
/// whose attributes or descendants are not completely represented by the
/// compact worksheet model is refused before any bytes are published.
pub(crate) fn replace_changed_rows(
    xml: &str,
    original: &[Sheet],
    candidate: &[Option<&Sheet>],
    max_output_bytes: usize,
) -> Result<String> {
    replace_changed_rows_impl(xml, original, candidate, max_output_bytes, false)?.ok_or_else(|| {
        invalid("flat ODS row transaction is not eligible for row-local publication")
    })
}

/// Try a row-local packaged worksheet publication without changing the
/// established structural-edit fallback.
pub(crate) fn try_replace_changed_rows(
    xml: &str,
    original: &[Sheet],
    candidate: &[Sheet],
    max_output_bytes: usize,
) -> Result<Option<String>> {
    if original.len() != candidate.len() {
        return Ok(None);
    }
    let changed = original
        .iter()
        .zip(candidate)
        .map(|(before, after)| (before != after).then_some(after))
        .collect::<Vec<_>>();
    replace_changed_rows_impl(xml, original, &changed, max_output_bytes, true)
}

fn replace_changed_rows_impl(
    xml: &str,
    original: &[Sheet],
    candidate: &[Option<&Sheet>],
    max_output_bytes: usize,
    allow_ineligible: bool,
) -> Result<Option<String>> {
    validation::validate_content_xml_size(xml)?;
    if candidate.len() > validation::MAX_PHYSICAL_RUNS {
        return Err(Error::InvalidFormat(format!(
            "flat ODS sheet count exceeds the {} safety limit",
            validation::MAX_PHYSICAL_RUNS
        )));
    }
    for sheet in candidate.iter().flatten() {
        validation::validate_sheet(sheet)?;
    }
    if original.len() != candidate.len() {
        if allow_ineligible {
            return Ok(None);
        }
        return Err(invalid(
            "flat ODS row transaction cannot add or remove worksheets",
        ));
    }

    let spans = scan(xml)?;
    let spreadsheet = one_spreadsheet(&spans)?;
    let mut tables = direct_children(&spans, spreadsheet, TABLE_NAMESPACE, "table");
    tables.sort_unstable_by_key(|index| spans[*index].start);
    if tables.len() != original.len() {
        return Err(invalid("flat ODS table inventory changed since parsing"));
    }
    if allow_ineligible {
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
    }

    let mut edits = Vec::new();
    for (sheet_index, ((before, after), table)) in
        original.iter().zip(candidate).zip(tables).enumerate()
    {
        let Some(after) = after else { continue };
        if before.name != after.name || before.style_name != after.style_name {
            if allow_ineligible {
                return Ok(None);
            }
            return Err(invalid("flat ODS row transaction cannot rename worksheets"));
        }
        let mut rows = direct_children(&spans, table, TABLE_NAMESPACE, "table-row");
        rows.sort_unstable_by_key(|index| spans[*index].start);
        if rows.len() != before.rows.len() {
            return Err(Error::InvalidFormat(format!(
                "flat ODS sheet {sheet_index} row inventory changed since parsing"
            )));
        }

        let prefix = before
            .rows
            .iter()
            .zip(&after.rows)
            .take_while(|(left, right)| left == right)
            .count();
        let suffix = before.rows[prefix..]
            .iter()
            .rev()
            .zip(after.rows[prefix..].iter().rev())
            .take_while(|(left, right)| left == right)
            .count();
        let old_end = before.rows.len() - suffix;
        let new_end = after.rows.len() - suffix;
        if prefix == old_end {
            if allow_ineligible {
                return Ok(None);
            }
            return Err(invalid(
                "flat ODS row insertion requires an existing physical row anchor",
            ));
        }

        for row in &rows[prefix..old_end] {
            validate_rewritable_row(xml, &spans, &spans[*row])?;
        }
        edits.push((
            spans[rows[prefix]].start,
            spans[rows[old_end - 1]].end,
            codec::write_rows_bounded(&after.rows[prefix..new_end], max_output_bytes)?,
        ));
    }

    apply_edits_bounded(xml, edits, max_output_bytes).map(Some)
}

fn apply_edits_bounded(
    xml: &str,
    mut edits: Vec<(usize, usize, String)>,
    max_output_bytes: usize,
) -> Result<String> {
    edits.sort_unstable_by_key(|(start, _, _)| *start);
    let mut cursor = 0usize;
    let mut removed = 0usize;
    let mut added = 0usize;
    for (start, end, replacement) in &edits {
        if *start < cursor || *end < *start || *end > xml.len() {
            return Err(invalid("overlapping or out-of-bounds flat ODS edit"));
        }
        removed = removed
            .checked_add(end - start)
            .ok_or_else(|| invalid("flat ODS removed-byte count overflow"))?;
        added = added
            .checked_add(replacement.len())
            .ok_or_else(|| invalid("flat ODS replacement-byte count overflow"))?;
        cursor = *end;
    }
    let output_len = xml
        .len()
        .checked_sub(removed)
        .and_then(|value| value.checked_add(added))
        .ok_or_else(|| invalid("flat ODS output size overflow"))?;
    if output_len > max_output_bytes {
        return Err(Error::InvalidFormat(format!(
            "flat ODS output exceeds the {max_output_bytes} byte limit"
        )));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_error| invalid("flat ODS output allocation failed"))?;
    cursor = 0;
    for (start, end, replacement) in edits {
        output.push_str(&xml[cursor..start]);
        output.push_str(&replacement);
        cursor = end;
    }
    output.push_str(&xml[cursor..]);
    Ok(output)
}

fn validate_rewritable_row(xml: &str, spans: &[Span], span: &Span) -> Result<()> {
    let row = xml
        .get(span.start..span.end)
        .ok_or_else(|| invalid("flat ODS row span is invalid"))?;
    let mut ancestors = Vec::new();
    let mut parent = span.parent;
    while let Some(index) = parent {
        let ancestor = spans
            .get(index)
            .ok_or_else(|| invalid("flat ODS row ancestor span is invalid"))?;
        if ancestor.empty {
            return Err(invalid("flat ODS row ancestor cannot be empty"));
        }
        ancestors.push(ancestor);
        parent = ancestor.parent;
    }
    if ancestors.is_empty() {
        return Err(invalid("flat ODS row has no document ancestors"));
    }
    let mut source = String::new();
    for ancestor in ancestors.iter().rev() {
        source.push_str(
            xml.get(ancestor.start..ancestor.tag_end)
                .ok_or_else(|| invalid("flat ODS row ancestor opening span is invalid"))?,
        );
    }
    source.push_str(row);
    for ancestor in &ancestors {
        source.push_str("</");
        source.push_str(&ancestor.qname);
        source.push('>');
    }
    let mut reader = NsReader::from_str(&source);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut row_depth = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid flat ODS row XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let namespace = resolve_namespace(&namespace)?;
                let local = decode(element.local_name().as_ref(), "row element local name")?;
                if row_depth == 0 {
                    if is_element_name(namespace.as_deref(), &local, TABLE_NAMESPACE, "table-row") {
                        validate_modeled_attributes(
                            &reader,
                            &element,
                            namespace.as_deref(),
                            &local,
                        )?;
                        row_depth = 1;
                    }
                    buffer.clear();
                    continue;
                }
                row_depth = row_depth.saturating_add(1);
                if !is_modeled_row_element(namespace.as_deref(), &local) {
                    return Err(Error::InvalidFormat(format!(
                        "flat ODS edit would discard unmodeled row element '{local}'"
                    )));
                }
                validate_modeled_attributes(&reader, &element, namespace.as_deref(), &local)?;
            },
            Event::Empty(element) => {
                let namespace = resolve_namespace(&namespace)?;
                let local = decode(element.local_name().as_ref(), "row element local name")?;
                if row_depth == 0 {
                    if is_element_name(namespace.as_deref(), &local, TABLE_NAMESPACE, "table-row") {
                        validate_modeled_attributes(
                            &reader,
                            &element,
                            namespace.as_deref(),
                            &local,
                        )?;
                    }
                    buffer.clear();
                    continue;
                }
                if !is_modeled_row_element(namespace.as_deref(), &local) {
                    return Err(Error::InvalidFormat(format!(
                        "flat ODS edit would discard unmodeled row element '{local}'"
                    )));
                }
                validate_modeled_attributes(&reader, &element, namespace.as_deref(), &local)?;
            },
            Event::Comment(_) | Event::PI(_) | Event::DocType(_) if row_depth > 0 => {
                return Err(invalid("flat ODS edit would discard unmodeled row markup"));
            },
            Event::End(_) if row_depth > 0 => row_depth -= 1,
            Event::Eof => break,
            Event::Decl(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::End(_) => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn is_element_name(
    namespace: Option<&str>,
    local: &str,
    expected_namespace: &str,
    expected_local: &str,
) -> bool {
    namespace == Some(expected_namespace) && local == expected_local
}

fn is_modeled_row_element(namespace: Option<&str>, local: &str) -> bool {
    (namespace == Some(TABLE_NAMESPACE)
        && matches!(local, "table-row" | "table-cell" | "covered-table-cell"))
        || (namespace == Some(codec::TEXT_NAMESPACE) && local == "p")
}

fn validate_modeled_attributes(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    element_namespace: Option<&str>,
    local: &str,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid flat ODS attribute: {error}"))
        })?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, attribute_local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolve_namespace(&namespace)?;
        let attribute_local = decode(attribute_local.as_ref(), "row attribute local name")?;
        let modeled = match (element_namespace, local, namespace.as_deref()) {
            (Some(TABLE_NAMESPACE), "table-row", Some(TABLE_NAMESPACE)) => matches!(
                attribute_local.as_str(),
                "number-rows-repeated" | "style-name" | "default-cell-style-name"
            ),
            (Some(TABLE_NAMESPACE), "table-cell" | "covered-table-cell", Some(TABLE_NAMESPACE)) => {
                matches!(
                    attribute_local.as_str(),
                    "number-columns-repeated"
                        | "number-rows-spanned"
                        | "number-columns-spanned"
                        | "formula"
                        | "style-name"
                )
            },
            (
                Some(TABLE_NAMESPACE),
                "table-cell" | "covered-table-cell",
                Some(OFFICE_NAMESPACE),
            ) => {
                matches!(
                    attribute_local.as_str(),
                    "value-type"
                        | "value"
                        | "date-value"
                        | "time-value"
                        | "boolean-value"
                        | "currency"
                )
            },
            (Some(codec::TEXT_NAMESPACE), "p", _) => false,
            _ => false,
        };
        if !modeled {
            return Err(Error::InvalidFormat(format!(
                "flat ODS edit would discard unmodeled attribute '{attribute_local}'"
            )));
        }
    }
    Ok(())
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
    let mut quote = None;
    for (index, byte) in xml.as_bytes()[..tag_end].iter().enumerate().rev() {
        match (quote, byte) {
            (Some(delimiter), current) if current == &delimiter => quote = None,
            (Some(_), _) => {},
            (None, b'\'' | b'"') => quote = Some(*byte),
            (None, b'<') => return Ok(index),
            _ => {},
        }
    }
    Err(Error::InvalidFormat(
        "ODS XML element start is missing".to_string(),
    ))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| Error::InvalidFormat("ODS XML position overflows usize".to_string()))
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
        .map_err(|_error| Error::InvalidFormat(format!("ODS {label} is not valid UTF-8")))
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
