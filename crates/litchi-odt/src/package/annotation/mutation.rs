//! Lossless XML edits and host/anchor validation for annotations.

use super::codec::{end_marker, serialize};
use super::model::{Annotation, AnnotationAnchor, AnnotationHost, AnnotationPosition, Scan, Site};
use super::scan::scan as scan_xml;
use super::{bounds, invalid, invalid_error};
use litchi_core::Result;

pub(crate) fn add_xml(
    content: &str,
    host: AnnotationHost,
    anchor: &AnnotationAnchor,
    annotation: &Annotation,
) -> Result<(String, usize)> {
    validate_anchor_host(host, anchor)?;
    annotation.validate()?;
    let scan = scan_xml(content, host)?;
    validate_new_name(&scan, annotation.name())?;
    if anchor.end.is_some() && annotation.name().is_none() {
        return invalid("a ranged annotation requires a non-empty office:name");
    }
    let start_site = site_for(&scan, &anchor.start)?;
    let start_at = insertion_position(start_site);
    let fragment = serialize(annotation)?;
    let updated = if let Some(end_position) = &anchor.end {
        let end_site = site_for(&scan, end_position)?;
        let end_at = insertion_position(end_site);
        if start_at > end_at {
            return invalid("annotation range end precedes its start");
        }
        let marker = end_marker(annotation.name().expect("range name validated"));
        if start_site.span.start == end_site.span.start {
            insert_child(content, start_site, &format!("{fragment}{marker}"))?
        } else {
            apply_edits(
                content,
                vec![
                    child_edit(content, start_site, &fragment)?,
                    child_edit(content, end_site, &marker)?,
                ],
            )?
        }
    } else {
        insert_child(content, start_site, &fragment)?
    };
    let index = scan
        .records
        .iter()
        .filter(|record| record.span.start < start_at)
        .count();
    scan_xml(&updated, host)?;
    Ok((updated, index))
}

pub(crate) fn replace_xml(
    content: &str,
    host: AnnotationHost,
    index: usize,
    annotation: &Annotation,
) -> Result<String> {
    annotation.validate()?;
    let scan = scan_xml(content, host)?;
    let record = scan
        .records
        .get(index)
        .ok_or_else(|| bounds(index, scan.records.len()))?;
    if record.end.is_some()
        && record.annotation.as_ref().and_then(Annotation::name) != annotation.name()
    {
        return invalid("replacing a ranged annotation cannot change its office:name");
    }
    if let Some(name) = annotation.name() {
        for (other_index, other) in scan.records.iter().enumerate() {
            if other_index != index
                && other.annotation.as_ref().and_then(Annotation::name) == Some(name)
            {
                return invalid(format!("duplicate annotation name '{name}'"));
            }
        }
    }
    let updated = apply_edits(
        content,
        vec![Edit {
            start: record.span.start,
            end: record.span.end,
            replacement: serialize(annotation)?,
        }],
    )?;
    scan_xml(&updated, host)?;
    Ok(updated)
}

pub(crate) fn remove_xml(content: &str, host: AnnotationHost, index: usize) -> Result<String> {
    let scan = scan_xml(content, host)?;
    let record = scan
        .records
        .get(index)
        .ok_or_else(|| bounds(index, scan.records.len()))?;
    let mut edits = vec![Edit {
        start: record.span.start,
        end: record.span.end,
        replacement: String::new(),
    }];
    if let Some((end, _)) = &record.end {
        edits.push(Edit {
            start: end.start,
            end: end.end,
            replacement: String::new(),
        });
    }
    let updated = apply_edits(content, edits)?;
    scan_xml(&updated, host)?;
    Ok(updated)
}

pub(crate) fn reorder_xml(
    content: &str,
    host: AnnotationHost,
    from: usize,
    to: usize,
) -> Result<String> {
    let scan = scan_xml(content, host)?;
    let first = scan
        .records
        .get(from)
        .ok_or_else(|| bounds(from, scan.records.len()))?;
    let second = scan
        .records
        .get(to)
        .ok_or_else(|| bounds(to, scan.records.len()))?;
    if first.end.is_some() || second.end.is_some() {
        return invalid("ranged annotations cannot be reordered independently of their text");
    }
    if first.parent_start != second.parent_start {
        return invalid("annotations can only be reordered among XML siblings");
    }
    if first.span.start == second.span.start {
        return Ok(content.to_string());
    }
    let (left, right) = if first.span.start < second.span.start {
        (first, second)
    } else {
        (second, first)
    };
    let mut updated = String::with_capacity(content.len());
    updated.push_str(&content[..left.span.start]);
    updated.push_str(&content[right.span.start..right.span.end]);
    updated.push_str(&content[left.span.end..right.span.start]);
    updated.push_str(&content[left.span.start..left.span.end]);
    updated.push_str(&content[right.span.end..]);
    scan_xml(&updated, host)?;
    Ok(updated)
}

pub(crate) fn validate_new_name(scan: &Scan, name: Option<&str>) -> Result<()> {
    let Some(name) = name else { return Ok(()) };
    if name.is_empty() {
        return invalid("annotation office:name cannot be empty");
    }
    if scan
        .records
        .iter()
        .any(|record| record.annotation.as_ref().and_then(Annotation::name) == Some(name))
    {
        return invalid(format!("duplicate annotation name '{name}'"));
    }
    Ok(())
}

pub(crate) fn validate_anchor_host(host: AnnotationHost, anchor: &AnnotationAnchor) -> Result<()> {
    let valid = |position: &AnnotationPosition| {
        matches!(
            (host, position),
            (_, AnnotationPosition::AnnotationBody { .. })
                | (
                    AnnotationHost::Text,
                    AnnotationPosition::TextParagraph { .. }
                )
                | (
                    AnnotationHost::Spreadsheet,
                    AnnotationPosition::SpreadsheetCell { .. }
                )
                | (
                    AnnotationHost::Presentation,
                    AnnotationPosition::PresentationPage { .. }
                )
                | (
                    AnnotationHost::Presentation,
                    AnnotationPosition::PresentationShape { .. }
                )
        )
    };
    if !valid(&anchor.start) || anchor.end.as_ref().is_some_and(|end| !valid(end)) {
        return invalid("annotation anchor does not belong to this document family");
    }
    if anchor.end.is_some()
        && host != AnnotationHost::Text
        && !matches!(anchor.start, AnnotationPosition::AnnotationBody { .. })
    {
        return invalid("named annotation ranges must be inserted in text paragraph content");
    }
    Ok(())
}

fn site_for<'a>(scan: &'a Scan, position: &AnnotationPosition) -> Result<&'a Site> {
    let mut matches = scan.sites.iter().filter(|site| &site.position == position);
    let site = matches
        .next()
        .ok_or_else(|| invalid_error(format!("annotation anchor {position:?} was not found")))?;
    if matches.next().is_some() && matches!(position, AnnotationPosition::PresentationShape { .. })
    {
        return invalid("presentation shape annotation anchor is ambiguous");
    }
    Ok(site)
}

fn insertion_position(site: &Site) -> usize {
    site.span
        .close_start
        .unwrap_or(site.span.end.saturating_sub(2))
}

fn insert_child(xml: &str, site: &Site, fragment: &str) -> Result<String> {
    apply_edits(xml, vec![child_edit(xml, site, fragment)?])
}

fn child_edit(xml: &str, site: &Site, fragment: &str) -> Result<Edit> {
    if let Some(close) = site.span.close_start {
        Ok(Edit {
            start: close,
            end: close,
            replacement: fragment.to_string(),
        })
    } else {
        let raw = xml
            .get(site.span.start..site.span.end)
            .ok_or_else(|| invalid_error("invalid empty annotation anchor span"))?;
        let slash = raw
            .rfind("/>")
            .ok_or_else(|| invalid_error("invalid empty annotation anchor"))?;
        Ok(Edit {
            start: site.span.start,
            end: site.span.end,
            replacement: format!("{}>{}</{}>", &raw[..slash], fragment, site.span.qname),
        })
    }
}

struct Edit {
    start: usize,
    end: usize,
    replacement: String,
}

fn apply_edits(xml: &str, mut edits: Vec<Edit>) -> Result<String> {
    edits.sort_by(|left, right| right.start.cmp(&left.start).then(right.end.cmp(&left.end)));
    let mut previous_start = xml.len();
    let mut output = xml.to_string();
    for edit in edits {
        if edit.start > edit.end || edit.end > xml.len() || edit.end > previous_start {
            return invalid("overlapping or invalid annotation XML edit");
        }
        output.replace_range(edit.start..edit.end, &edit.replacement);
        previous_start = edit.start;
    }
    Ok(output)
}
