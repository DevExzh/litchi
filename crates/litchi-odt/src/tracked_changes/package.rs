//! Document-level tracked-change declaration and marker mutations.

use super::codec::{
    apply_tracked_edits, escaped_tracked_id, resolve_story_position, scan_mutable_tracked_xml,
    validate_authored_tracked_xml,
};
use super::{Position, invalid, make_error};
use crate::parser::{ChangeType, Parser, TrackedChanges};
use litchi_core::Result;

/// Install or remove the declaration table without rewriting unrelated XML.
pub fn set_tracked_changes_xml(xml: &str, tracked: Option<&TrackedChanges>) -> Result<String> {
    let sites = scan_mutable_tracked_xml(xml)?;
    let fragment = tracked.map(TrackedChanges::to_xml_fragment).transpose()?;
    let output = match (sites.tracked_changes, fragment) {
        (Some(span), Some(fragment)) => apply_tracked_edits(xml, vec![(span, fragment)])?,
        (Some(span), None) => apply_tracked_edits(xml, vec![(span, String::new())])?,
        (None, Some(fragment)) => apply_tracked_edits(
            xml,
            vec![(
                sites.office_text_open_end..sites.office_text_open_end,
                fragment,
            )],
        )?,
        (None, None) => xml.to_string(),
    };
    validate_authored_tracked_xml(&output)?;
    Ok(output)
}

/// Insert a start/end marker pair at Unicode-safe story positions.
pub fn mark_tracked_change_range_xml(
    xml: &str,
    change_id: &str,
    start: &Position,
    end: &Position,
) -> Result<String> {
    let tracked = Parser::parse_tracked_changes(xml)?;
    let change = tracked
        .changes
        .iter()
        .find(|change| change.id == change_id)
        .ok_or_else(|| make_error(format!("unknown tracked-change ID '{change_id}'")))?;
    if change.change_type == ChangeType::Deletion {
        return invalid("deletion declarations require a point text:change marker");
    }
    let sites = scan_mutable_tracked_xml(xml)?;
    let start_offset = resolve_story_position(&sites.stories, start)?;
    let end_offset = resolve_story_position(&sites.stories, end)?;
    if start_offset >= end_offset {
        return invalid("tracked-change range start must precede its end");
    }
    let id = escaped_tracked_id(change_id);
    let start_fragment = format!(
        r#"<text:change-start xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" text:change-id="{id}"/>"#
    );
    let end_fragment = format!(
        r#"<text:change-end xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" text:change-id="{id}"/>"#
    );
    let output = apply_tracked_edits(
        xml,
        vec![
            (end_offset..end_offset, end_fragment),
            (start_offset..start_offset, start_fragment),
        ],
    )?;
    validate_authored_tracked_xml(&output)?;
    Ok(output)
}

/// Insert a deletion point marker at a Unicode-safe story position.
pub fn mark_tracked_deletion_xml(
    xml: &str,
    change_id: &str,
    position: &Position,
) -> Result<String> {
    let tracked = Parser::parse_tracked_changes(xml)?;
    let change = tracked
        .changes
        .iter()
        .find(|change| change.id == change_id)
        .ok_or_else(|| make_error(format!("unknown tracked-change ID '{change_id}'")))?;
    if change.change_type != ChangeType::Deletion {
        return invalid("point text:change markers require a deletion declaration");
    }
    let sites = scan_mutable_tracked_xml(xml)?;
    let site = sites
        .stories
        .iter()
        .find(|site| site.story == position.story)
        .ok_or_else(|| make_error("tracked-change story was not found"))?;
    let id = escaped_tracked_id(change_id);
    let fragment = format!(
        r#"<text:change xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" text:change-id="{id}"/>"#
    );
    let output = if let Some((span, qname)) = &site.empty {
        if position.character != 0 {
            return invalid("tracked-change character offset is out of bounds");
        }
        let source = xml
            .get(span.clone())
            .ok_or_else(|| make_error("invalid empty story span"))?;
        let open = source
            .strip_suffix("/>")
            .ok_or_else(|| make_error("empty story does not end with />"))?;
        apply_tracked_edits(
            xml,
            vec![(span.clone(), format!("{open}>{fragment}</{qname}>"))],
        )?
    } else {
        let offset = resolve_story_position(&sites.stories, position)?;
        apply_tracked_edits(xml, vec![(offset..offset, fragment)])?
    };
    validate_authored_tracked_xml(&output)?;
    Ok(output)
}

/// Remove every marker referencing one declaration while retaining its live text.
pub fn unmark_tracked_change_xml(xml: &str, change_id: &str) -> Result<String> {
    let sites = scan_mutable_tracked_xml(xml)?;
    let edits = sites
        .markers
        .into_iter()
        .filter(|marker| marker.id == change_id)
        .map(|marker| (marker.span, String::new()))
        .collect();
    let output = apply_tracked_edits(xml, edits)?;
    validate_authored_tracked_xml(&output)?;
    Ok(output)
}
