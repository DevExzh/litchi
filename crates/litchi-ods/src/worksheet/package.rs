//! Package-boundary replacement of direct spreadsheet tables.
//!
//! The package layer edits only the table spans owned by
//! `office:spreadsheet`.  Unmodelled spreadsheet children make a semantic
//! replacement unsafe, so the operation rejects them instead of silently
//! discarding producer data.

use super::{CellValue, Sheet, codec, validation};
use litchi_core::{Error, Result};
use litchi_odf_common::{
    constants,
    core::{AuthoredXmlFragment, OwnedPackage, XmlSourcePart, XmlSplicePublication},
};
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

#[derive(Clone, Debug)]
struct RowEdit {
    start: usize,
    end: usize,
    replacement: String,
}

/// One row-local result retaining both assembled content and its exact source
/// range proofs for package publication.
pub(crate) struct ChangedRows {
    pub(crate) content: String,
    pub(crate) publication: XmlSplicePublication,
}

/// Replace the direct table children of `office:spreadsheet` in one bounded
/// XML pass.  The returned content is structurally and semantically rechecked
/// before publication by the caller.
pub(crate) fn replace_tables(xml: &str, sheets: &[Sheet]) -> Result<String> {
    replace_tables_bounded(xml, sheets, validation::MAX_CONTENT_XML_BYTES)
}

fn replace_tables_bounded(xml: &str, sheets: &[Sheet], max_output_bytes: usize) -> Result<String> {
    validation::validate_content_xml_size(xml)?;
    validation::validate_sheets(sheets)?;
    let spans = scan(xml)?;
    let spreadsheet = one_spreadsheet(&spans)?;
    let mut tables = direct_children(&spans, spreadsheet, TABLE_NAMESPACE, "table");
    tables.sort_unstable_by_key(|index| spans[*index].start);

    for table in &tables {
        validate_rewritable_table_content(xml, &spans, *table)?;
        for child in direct_children_any(&spans, *table) {
            if !is_element(&spans[child], TABLE_NAMESPACE, "table-row") {
                return Err(Error::InvalidFormat(
                    "ODS worksheet edit cannot replace a table containing unmodeled direct children"
                        .to_string(),
                ));
            }
        }
        for row in direct_children(&spans, *table, TABLE_NAMESPACE, "table-row") {
            validate_rewritable_row(xml, &spans, row)?;
        }
    }

    let removed_bytes = tables.iter().try_fold(0usize, |total, table| {
        total
            .checked_add(spans[*table].end - spans[*table].start)
            .ok_or_else(|| invalid("flat ODS removed table size overflows usize"))
    })?;
    let mut retained_bytes = xml
        .len()
        .checked_sub(removed_bytes)
        .ok_or_else(|| invalid("flat ODS retained table size underflows usize"))?;
    let expands_empty_spreadsheet = sheets.len() > tables.len() && spans[spreadsheet].empty;
    let empty_opening = if expands_empty_spreadsheet {
        let opening = xml
            .get(spans[spreadsheet].start..spans[spreadsheet].tag_end)
            .ok_or_else(|| invalid("invalid spreadsheet start span"))?;
        let opening = opening
            .trim_end()
            .strip_suffix("/>")
            .ok_or_else(|| invalid("empty spreadsheet has no self-closing tag"))?;
        let expanded_shell_bytes = opening
            .len()
            .checked_add(spans[spreadsheet].qname.len())
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| invalid("flat ODS spreadsheet expansion size overflows usize"))?;
        retained_bytes = retained_bytes
            .checked_sub(spans[spreadsheet].tag_end - spans[spreadsheet].start)
            .and_then(|value| value.checked_add(expanded_shell_bytes))
            .ok_or_else(|| invalid("flat ODS spreadsheet expansion size overflows usize"))?;
        Some(opening.to_string())
    } else {
        None
    };
    if retained_bytes > max_output_bytes {
        return Err(Error::InvalidFormat(format!(
            "flat ODS output exceeds the {max_output_bytes} byte limit"
        )));
    }

    let mut edits = Vec::<RowEdit>::with_capacity(tables.len().max(sheets.len()).saturating_add(1));
    let mut rendered_bytes = 0usize;
    let mut extras = String::new();
    for (position, sheet) in sheets.iter().enumerate() {
        let remaining = max_output_bytes
            .checked_sub(retained_bytes)
            .and_then(|value| value.checked_sub(rendered_bytes))
            .ok_or_else(|| invalid("flat ODS worksheet output budget underflows usize"))?;
        let replacement = codec::write_sheet_bounded(sheet, remaining)?;
        rendered_bytes = rendered_bytes
            .checked_add(replacement.len())
            .ok_or_else(|| invalid("flat ODS rendered table size overflows usize"))?;
        if let Some(table) = tables.get(position) {
            edits.push(RowEdit {
                start: spans[*table].start,
                end: spans[*table].end,
                replacement,
            });
        } else {
            extras
                .try_reserve_exact(replacement.len())
                .map_err(|_error| invalid("flat ODS extra table allocation failed"))?;
            extras.push_str(&replacement);
        }
    }
    for table in tables.iter().skip(sheets.len()) {
        edits.push(RowEdit {
            start: spans[*table].start,
            end: spans[*table].end,
            replacement: String::new(),
        });
    }

    if !extras.is_empty() {
        if spans[spreadsheet].empty {
            let opening = empty_opening
                .as_deref()
                .ok_or_else(|| invalid("empty spreadsheet opening is missing"))?;
            let prefix = format!("{opening}>");
            let suffix = format!("</{}>", spans[spreadsheet].qname);
            extras
                .try_reserve_exact(prefix.len().saturating_add(suffix.len()))
                .map_err(|_error| invalid("flat ODS spreadsheet expansion allocation failed"))?;
            extras.insert_str(0, &prefix);
            extras.push_str(&suffix);
            edits.push(RowEdit {
                start: spans[spreadsheet].start,
                end: spans[spreadsheet].tag_end,
                replacement: extras,
            });
        } else {
            let insertion = direct_children_any(&spans, spreadsheet)
                .into_iter()
                .filter(|child| !is_element(&spans[*child], TABLE_NAMESPACE, "table"))
                .map(|child| spans[child].start)
                .min()
                .unwrap_or(spans[spreadsheet].close_start);
            edits.push(RowEdit {
                start: insertion,
                end: insertion,
                replacement: extras,
            });
        }
    }

    let updated = apply_edits_bounded(xml, edits, max_output_bytes)?;
    crate::authoring::validate_content_xml(&updated)?;
    codec::parse(&updated)?;
    Ok(updated)
}

fn validate_rewritable_table_content(xml: &str, spans: &[Span], span_index: usize) -> Result<()> {
    let table = spans
        .get(span_index)
        .ok_or_else(|| invalid("flat ODS table span is invalid"))?;
    validate_rewritable_table_attributes(xml, spans, span_index)?;
    if table.empty {
        return Ok(());
    }
    let mut children = direct_children_any(spans, span_index);
    children.sort_unstable_by_key(|child| spans[*child].start);
    let mut cursor = table.tag_end;
    for child in children {
        validate_ignorable_table_gap(xml, cursor, spans[child].start)?;
        cursor = spans[child].end;
    }
    validate_ignorable_table_gap(xml, cursor, table.close_start)
}

fn validate_rewritable_table_attributes(
    xml: &str,
    spans: &[Span],
    span_index: usize,
) -> Result<()> {
    let table = spans
        .get(span_index)
        .ok_or_else(|| invalid("flat ODS table span is invalid"))?;
    let mut ancestors = Vec::new();
    let mut parent = table.parent;
    while let Some(index) = parent {
        let ancestor = spans
            .get(index)
            .ok_or_else(|| invalid("flat ODS table ancestor span is invalid"))?;
        if ancestor.empty {
            return Err(invalid("flat ODS table ancestor cannot be empty"));
        }
        ancestors.push(ancestor);
        parent = ancestor.parent;
    }
    if ancestors.is_empty() {
        return Err(invalid("flat ODS table has no document ancestors"));
    }

    let mut source = String::new();
    for ancestor in ancestors.iter().rev() {
        source.push_str(
            xml.get(ancestor.start..ancestor.tag_end)
                .ok_or_else(|| invalid("flat ODS table ancestor opening span is invalid"))?,
        );
    }
    let opening = xml
        .get(table.start..table.tag_end)
        .ok_or_else(|| invalid("flat ODS table opening span is invalid"))?;
    if table.empty {
        source.push_str(opening);
    } else {
        source.push_str(
            opening
                .strip_suffix('>')
                .ok_or_else(|| invalid("flat ODS table opening tag is invalid"))?,
        );
        source.push_str("/>");
    }
    for ancestor in &ancestors {
        source.push_str("</");
        source.push_str(&ancestor.qname);
        source.push('>');
    }

    let mut reader = NsReader::from_str(&source);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid flat ODS table XML: {error}"))
            })?;
        match event {
            Event::Empty(element) => {
                let namespace = resolve_namespace(&namespace)?;
                let local = decode(element.local_name().as_ref(), "table element local name")?;
                if is_element_name(namespace.as_deref(), &local, TABLE_NAMESPACE, "table") {
                    for attribute in element.attributes().with_checks(true) {
                        let attribute = attribute.map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid flat ODS table attribute: {error}"
                            ))
                        })?;
                        let name = attribute.key.as_ref();
                        if name == b"xmlns" || name.starts_with(b"xmlns:") {
                            continue;
                        }
                        let (attribute_namespace, attribute_local) =
                            reader.resolver().resolve_attribute(attribute.key);
                        let attribute_namespace = resolve_namespace(&attribute_namespace)?;
                        let attribute_local =
                            decode(attribute_local.as_ref(), "table attribute local name")?;
                        if attribute_namespace.as_deref() != Some(TABLE_NAMESPACE)
                            || !matches!(attribute_local.as_str(), "name" | "style-name")
                        {
                            return Err(Error::InvalidFormat(format!(
                                "flat ODS edit would discard unmodeled table attribute '{attribute_local}'"
                            )));
                        }
                    }
                    return Ok(());
                }
            },
            Event::Eof => return Err(invalid("flat ODS table opening element is missing")),
            _ => {},
        }
        buffer.clear();
    }
}

fn validate_ignorable_table_gap(xml: &str, start: usize, end: usize) -> Result<()> {
    let gap = xml
        .get(start..end)
        .ok_or_else(|| invalid("flat ODS table content span is invalid"))?;
    if gap
        .bytes()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
    {
        return Ok(());
    }
    Err(invalid(
        "flat ODS edit would discard unmodeled table-level markup",
    ))
}

/// The scanned element layout of one immutable `content.xml` projection.
///
/// Spans are owned offsets into the scanned XML, and `tables`/`rows` are the
/// spreadsheet's direct table children and each table's sorted direct
/// table-row children, derived deterministically from the spans. A layout
/// computed for one document must only be reused with byte-identical input.
/// The `SourceBackedSpreadsheet` owner satisfies this by construction: its
/// retained `content_xml` never changes for the owner's lifetime.
#[derive(Debug)]
pub(crate) struct ContentLayout {
    spans: Vec<Span>,
    tables: Vec<usize>,
    rows: Vec<Vec<usize>>,
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
    let edits = changed_row_edits(xml, None, original, candidate, max_output_bytes, false)?
        .0
        .ok_or_else(|| {
            invalid("flat ODS row transaction is not eligible for row-local publication")
        })?;
    apply_edits_bounded(xml, edits, max_output_bytes)
}

/// Rewrite only the changed physical rows of a source-backed package using a
/// previously scanned layout of the same `xml`.
///
/// The layout must come from a successful scan of byte-identical `xml`
/// (see [`replace_changed_rows_from_content_xml_retaining_layout`]); the
/// per-call gates still run before the layout is consulted, in the same order
/// as the scanning variant.
pub(crate) fn replace_changed_rows_from_content_xml_with_layout(
    xml: &str,
    layout: &ContentLayout,
    original: &[Sheet],
    candidate: &[Option<&Sheet>],
    max_output_bytes: usize,
) -> Result<Option<String>> {
    let (edits, _layout) = changed_row_edits(
        xml,
        Some(layout),
        original,
        candidate,
        max_output_bytes,
        true,
    )?;
    edits
        .map(|edits| apply_edits_bounded(xml, edits, max_output_bytes))
        .transpose()
}

/// Rewrite only the changed physical rows of a source-backed package,
/// retaining the scanned layout so the caller can cache it for later
/// transactions over the same immutable `xml`.
///
/// The layout is returned only when the scan actually ran and succeeded; a
/// gate refusal before the scan returns `(None, None)` and a scan failure
/// propagates the error without a layout, so errors are never cached.
pub(crate) fn replace_changed_rows_from_content_xml_retaining_layout(
    xml: &str,
    original: &[Sheet],
    candidate: &[Option<&Sheet>],
    max_output_bytes: usize,
) -> Result<(Option<String>, Option<ContentLayout>)> {
    let (edits, layout) =
        changed_row_edits(xml, None, original, candidate, max_output_bytes, true)?;
    let content = edits
        .map(|edits| apply_edits_bounded(xml, edits, max_output_bytes))
        .transpose()?;
    Ok((content, layout))
}

/// Try a provenance-bearing row-local packaged worksheet publication without
/// changing the established structural-edit fallback.
pub(crate) fn try_replace_changed_rows_spliced(
    source: &OwnedPackage,
    original: &[Sheet],
    candidate: &[Sheet],
    max_output_bytes: usize,
) -> Result<Option<ChangedRows>> {
    let source_bytes = source.get_file(constants::ODF_CONTENT)?;
    let xml = std::str::from_utf8(&source_bytes)
        .map_err(|error| Error::InvalidFormat(format!("invalid ODS content.xml UTF-8: {error}")))?;
    if original.len() != candidate.len() {
        return Ok(None);
    }
    let changed = original
        .iter()
        .zip(candidate)
        .map(|(before, after)| (before != after).then_some(after))
        .collect::<Vec<_>>();
    let (Some(edits), _layout) =
        changed_row_edits(xml, None, original, &changed, max_output_bytes, true)?
    else {
        return Ok(None);
    };
    let content = apply_edits_bounded(xml, edits.clone(), max_output_bytes)?;
    let source_part = XmlSourcePart::load(source, constants::ODF_CONTENT)?;
    let mut publication = XmlSplicePublication::new(source_part.clone());
    for edit in edits {
        let expected = source_part
            .bytes()
            .get(edit.start..edit.end)
            .ok_or_else(|| {
                Error::InvalidFormat("ODS row splice source range is invalid".to_string())
            })?;
        let proof = source_part.checked_range(edit.start..edit.end, expected)?;
        let fragment = if edit.replacement.is_empty() {
            AuthoredXmlFragment::deletion()
        } else {
            AuthoredXmlFragment::markup(edit.replacement.into_bytes())?
        };
        publication.replace(proof, fragment)?;
    }
    Ok(Some(ChangedRows {
        content,
        publication,
    }))
}

fn changed_row_edits(
    xml: &str,
    layout: Option<&ContentLayout>,
    original: &[Sheet],
    candidate: &[Option<&Sheet>],
    max_output_bytes: usize,
    allow_ineligible: bool,
) -> Result<(Option<Vec<RowEdit>>, Option<ContentLayout>)> {
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
            return Ok((None, None));
        }
        return Err(invalid(
            "flat ODS row transaction cannot add or remove worksheets",
        ));
    }

    // Scan and derive the table/row topology only when no cached layout was
    // supplied; the gates above always run first, preserving the established
    // error ordering on both paths. `owned_layout` stays empty (no
    // allocation) when a cached layout is used.
    let mut owned_layout = None;
    let layout: &ContentLayout = match layout {
        Some(layout) => layout,
        None => {
            owned_layout = Some(build_layout(scan(xml)?)?);
            owned_layout
                .as_ref()
                .ok_or_else(|| invalid("flat ODS layout is missing"))?
        },
    };
    let spans: &[Span] = &layout.spans;
    if layout.tables.len() != original.len() {
        return Err(invalid("flat ODS table inventory changed since parsing"));
    }
    let mut edits = Vec::new();
    for (sheet_index, ((before, after), rows)) in
        original.iter().zip(candidate).zip(&layout.rows).enumerate()
    {
        let Some(after) = after else { continue };
        if before.name != after.name || before.style_name != after.style_name {
            if allow_ineligible {
                return Ok((None, None));
            }
            return Err(invalid("flat ODS row transaction cannot rename worksheets"));
        }
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
                return Ok((None, None));
            }
            return Err(invalid(
                "flat ODS row insertion requires an existing physical row anchor",
            ));
        }

        if allow_ineligible
            && before.rows[prefix..old_end]
                .iter()
                .flat_map(|row| &row.cells)
                .any(|cell| matches!(cell.value, CellValue::Unknown { .. }))
        {
            return Err(invalid(
                "ODS source row span contains an unknown value that cannot be regenerated",
            ));
        }

        for row in &rows[prefix..old_end] {
            validate_rewritable_row(xml, spans, *row)?;
        }
        edits.push(RowEdit {
            start: spans[rows[prefix]].start,
            end: spans[rows[old_end - 1]].end,
            replacement: codec::write_rows_bounded(&after.rows[prefix..new_end], max_output_bytes)?,
        });
    }

    // `owned_layout` is `Some` exactly when this call ran the scan; cached
    // calls return `None`, matching the established retention contract.
    Ok((Some(edits), owned_layout))
}

fn build_layout(spans: Vec<Span>) -> Result<ContentLayout> {
    let spreadsheet = one_spreadsheet(&spans)?;
    let mut tables = direct_children(&spans, spreadsheet, TABLE_NAMESPACE, "table");
    tables.sort_unstable_by_key(|index| spans[*index].start);
    let mut rows = Vec::with_capacity(tables.len());
    for table in &tables {
        let mut table_rows = direct_children(&spans, *table, TABLE_NAMESPACE, "table-row");
        table_rows.sort_unstable_by_key(|index| spans[*index].start);
        rows.push(table_rows);
    }
    Ok(ContentLayout {
        spans,
        tables,
        rows,
    })
}

fn apply_edits_bounded(
    xml: &str,
    mut edits: Vec<RowEdit>,
    max_output_bytes: usize,
) -> Result<String> {
    edits.sort_unstable_by_key(|edit| edit.start);
    let mut cursor = 0usize;
    let mut removed = 0usize;
    let mut added = 0usize;
    for edit in &edits {
        if edit.start < cursor || edit.end < edit.start || edit.end > xml.len() {
            return Err(invalid("overlapping or out-of-bounds flat ODS edit"));
        }
        removed = removed
            .checked_add(edit.end - edit.start)
            .ok_or_else(|| invalid("flat ODS removed-byte count overflow"))?;
        added = added
            .checked_add(edit.replacement.len())
            .ok_or_else(|| invalid("flat ODS replacement-byte count overflow"))?;
        cursor = edit.end;
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
    for edit in edits {
        output.push_str(&xml[cursor..edit.start]);
        output.push_str(&edit.replacement);
        cursor = edit.end;
    }
    output.push_str(&xml[cursor..]);
    Ok(output)
}

fn validate_rewritable_row(xml: &str, spans: &[Span], span_index: usize) -> Result<()> {
    let span = spans
        .get(span_index)
        .ok_or_else(|| invalid("flat ODS row span is invalid"))?;
    let mut paragraph_owner = None;
    for child in spans
        .iter()
        .skip(span_index.saturating_add(1))
        .take_while(|child| child.start < span.end)
    {
        if child.parent == Some(span_index) {
            if is_element(child, TABLE_NAMESPACE, "table-cell")
                || is_element(child, TABLE_NAMESPACE, "covered-table-cell")
            {
                continue;
            }
            return Err(Error::InvalidFormat(format!(
                "flat ODS edit would discard unmodeled row element '{}'",
                child.local
            )));
        }
        let parent_index = child
            .parent
            .ok_or_else(|| invalid("flat ODS row descendant has no parent"))?;
        let parent = spans
            .get(parent_index)
            .ok_or_else(|| invalid("flat ODS row descendant parent is invalid"))?;
        if parent.parent != Some(span_index)
            || !is_element(child, codec::TEXT_NAMESPACE, "p")
            || paragraph_owner == Some(parent_index)
        {
            if child.namespace.as_deref() == Some(codec::TEXT_NAMESPACE) && child.local != "p" {
                return Err(invalid(&format!(
                    "flat ODS edit would discard unsupported inline content '{}'",
                    child.qname
                )));
            }
            return Err(invalid(
                "flat ODS edit requires at most one direct text paragraph per cell",
            ));
        }
        paragraph_owner = Some(parent_index);
    }
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
    let mut in_paragraph = false;
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
                if is_element_name(namespace.as_deref(), &local, codec::TEXT_NAMESPACE, "p") {
                    in_paragraph = true;
                }
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
            Event::Text(text) if row_depth > 0 => {
                let bytes: &[u8] = text.as_ref();
                if !in_paragraph
                    && !bytes
                        .iter()
                        .all(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
                {
                    return Err(invalid(
                        "flat ODS edit would discard text outside a cell paragraph",
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if row_depth > 0 => {
                return Err(invalid(
                    "flat ODS edit would discard unsupported cell text markup",
                ));
            },
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) | Event::DocType(_)
                if row_depth > 0 =>
            {
                return Err(invalid("flat ODS edit would discard unmodeled row markup"));
            },
            Event::End(element) if row_depth > 0 => {
                let local = decode(element.local_name().as_ref(), "row element local name")?;
                if local == "p" && in_paragraph {
                    in_paragraph = false;
                }
                row_depth -= 1;
            },
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
        let event_start = position(&reader)?;
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
                    xml,
                    &mut spans,
                    &reader,
                    namespace,
                    local,
                    qname,
                    parent,
                    false,
                    event_start,
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
                    xml,
                    &mut spans,
                    &reader,
                    namespace,
                    local,
                    qname,
                    parent,
                    true,
                    event_start,
                )?;
            },
            Event::End(_) => {
                let index = open.pop().ok_or_else(|| {
                    Error::InvalidFormat("ODS XML element stack underflow".to_string())
                })?;
                let end = position(&reader)?;
                if xml
                    .as_bytes()
                    .get(event_start..event_start.saturating_add(2))
                    != Some(b"</")
                {
                    return Err(Error::InvalidFormat(
                        "ODS XML closing tag start is missing".to_string(),
                    ));
                }
                spans[index].close_start = event_start;
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
    start: usize,
) -> Result<usize> {
    if spans.len() >= MAX_SPANS {
        return Err(Error::InvalidFormat(
            "ODS content contains too many XML elements".to_string(),
        ));
    }
    let tag_end = position(reader)?;
    if xml.as_bytes().get(start) != Some(&b'<') || start >= tag_end {
        return Err(Error::InvalidFormat(
            "ODS XML element start is missing".to_string(),
        ));
    }
    let index = spans.len();
    spans.push(Span {
        namespace,
        local,
        qname,
        start,
        tag_end,
        close_start: tag_end,
        end: tag_end,
        parent,
        empty,
    });
    Ok(index)
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

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_string())
}

#[cfg(test)]
mod fallback_bound_tests {
    use super::replace_tables_bounded;
    use crate::worksheet::{Sheet, codec};

    const EMPTY_CONTENT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#,
        r#"<office:body><office:spreadsheet/></office:body></office:document-content>"#,
    );

    #[test]
    fn cumulative_sheet_budget_rejects_before_combined_publication() {
        let sheets = [
            Sheet::new("First").expect("test sheet should be valid"),
            Sheet::new("Second").expect("test sheet should be valid"),
        ];
        let complete = replace_tables_bounded(EMPTY_CONTENT, &sheets, usize::MAX)
            .expect("small replacement should render");
        let limit = complete.len() - 1;
        assert!(codec::write_sheet_bounded(&sheets[0], limit).is_ok());
        assert!(codec::write_sheet_bounded(&sheets[1], limit).is_ok());

        let error = replace_tables_bounded(EMPTY_CONTENT, &sheets, limit)
            .expect_err("cumulative output must honor the shared byte budget");
        assert!(error.to_string().contains("byte limit"));
    }
}
