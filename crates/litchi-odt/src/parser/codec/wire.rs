//! Namespace-aware XML wire codecs for ODT-specific document structures.

use super::super::MAX_SEMANTIC_ITEMS;
use super::super::model::{
    ChangeType, Comment, Section, SectionDdeSource, SectionDisplay, SectionSource, TrackChange,
    TrackedChanges,
};
use super::semantic::{
    ActiveComment, ActiveSection, ActiveTrackedChange, ChangeRangeState, PendingAnnotation,
    PendingChangeRange,
};
use super::validation::{
    checked_semantic_depth, ensure_pending_capacity, parse_boolean, parse_tracked_change_bool,
    validate_protection_key, validate_tracked_change_text,
};
use crate::elements::xml::{
    DC_NAMESPACE, META_NAMESPACE, OFFICE_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE, XML_NAMESPACE,
    append_checked, append_text_control, decode_reference, is_bound, namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

pub(crate) fn parse_change_declarations(content: &str) -> Result<TrackedChanges> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut tracked_depth = 0usize;
    let mut tracked_changes_seen = false;
    let mut active: Option<ActiveTrackedChange> = None;
    let mut changes = Vec::new();
    let mut ids = HashMap::new();
    let mut tracked = TrackedChanges::default();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid tracked-change XML: {error}"))
            })?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        let dc_element = is_bound(&namespace, DC_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "tracked change")?;
                if tracked_depth == 0 {
                    if text_element && element.local_name().as_ref() == b"tracked-changes" {
                        if tracked_changes_seen {
                            return Err(Error::InvalidFormat(
                                "multiple text:tracked-changes elements are not allowed"
                                    .to_string(),
                            ));
                        }
                        tracked_changes_seen = true;
                        parse_tracked_changes_attributes(&reader, element, &mut tracked)?;
                        tracked_depth = 1;
                    }
                } else {
                    tracked_depth = checked_semantic_depth(tracked_depth, "tracked change")?;
                    if let Some(change) = active.as_mut() {
                        change.depth = checked_semantic_depth(change.depth, "changed region")?;
                        process_change_declaration_start(
                            &reader,
                            element,
                            text_element,
                            office_element,
                            dc_element,
                            change,
                        )?;
                    } else if text_element && element.local_name().as_ref() == b"changed-region" {
                        if tracked_depth != 2 {
                            return Err(Error::InvalidFormat(
                                "text:changed-region must be a direct child of text:tracked-changes"
                                    .to_string(),
                            ));
                        }
                        if changes.len() >= MAX_SEMANTIC_ITEMS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_SEMANTIC_ITEMS} tracked changes"
                            )));
                        }
                        let (id, xml_id) = change_region_ids(&reader, element)?;
                        if ids.insert(id.clone(), changes.len()).is_some() {
                            return Err(Error::InvalidFormat(format!(
                                "duplicate tracked-change ID '{id}'"
                            )));
                        }
                        active = Some(ActiveTrackedChange::new(id, xml_id));
                    }
                }
            },
            Event::Empty(ref element) if tracked_depth > 0 => {
                if text_element && element.local_name().as_ref() == b"changed-region" {
                    return Err(Error::InvalidFormat(
                        "text:changed-region requires a change declaration".to_string(),
                    ));
                }
                if let Some(change) = active.as_mut() {
                    process_change_declaration_empty(
                        &reader,
                        element,
                        text_element,
                        office_element,
                        dc_element,
                        change,
                    )?;
                }
            },
            Event::Empty(ref element)
                if text_element && element.local_name().as_ref() == b"tracked-changes" =>
            {
                if tracked_changes_seen {
                    return Err(Error::InvalidFormat(
                        "multiple text:tracked-changes elements are not allowed".to_string(),
                    ));
                }
                tracked_changes_seen = true;
                parse_tracked_changes_attributes(&reader, element, &mut tracked)?;
            },
            Event::Text(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid tracked-change text: {error}"))
                    })?;
                append_change_declaration_text(
                    active.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("missing active tracked change".to_string())
                    })?,
                    &value,
                )?;
            },
            Event::CData(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid tracked-change CDATA: {error}"))
                    })?;
                append_change_declaration_text(
                    active.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("missing active tracked change".to_string())
                    })?,
                    &value,
                )?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let value = decode_reference(reference, "tracked change")?;
                append_change_declaration_text(
                    active.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("missing active tracked change".to_string())
                    })?,
                    &value,
                )?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("tracked-change XML stack underflow".to_string())
                })?;
                if tracked_depth > 0 {
                    if let Some(change) = active.as_mut() {
                        if change.creator_depth == Some(change.depth) {
                            change.creator_depth = None;
                        }
                        if change.date_depth == Some(change.depth) {
                            change.date_depth = None;
                        }
                        if change.comment_depth == Some(change.depth) {
                            change.comment_depth = None;
                        }
                        if change.paragraph_depth == Some(change.depth) {
                            change.paragraph_depth = None;
                        }
                        if change.change_info_depth == Some(change.depth) {
                            change.change_info_depth = None;
                        }
                        if change.kind_depth == Some(change.depth) {
                            change.kind_depth = None;
                        }
                        change.depth = change.depth.checked_sub(1).ok_or_else(|| {
                            Error::InvalidFormat("changed-region stack underflow".to_string())
                        })?;
                        if change.depth == 0 {
                            changes.push(
                                active
                                    .take()
                                    .ok_or_else(|| {
                                        Error::InvalidFormat(
                                            "missing completed tracked change".to_string(),
                                        )
                                    })?
                                    .finish()?,
                            );
                        }
                    }
                    tracked_depth = tracked_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("tracked-changes stack underflow".to_string())
                    })?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || tracked_depth != 0 || active.is_some() {
        return Err(Error::InvalidFormat(
            "incomplete tracked-change XML structure".to_string(),
        ));
    }
    tracked.changes = changes;
    Ok(tracked)
}

fn change_region_ids(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<(String, Option<String>)> {
    let text_id = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"id", "changed-region")?;
    let xml_id = namespaced_attribute(reader, element, XML_NAMESPACE, b"id", "changed-region")?;
    let (id, xml_id) = match (text_id, xml_id) {
        (Some(text_id), Some(xml_id)) => {
            if text_id != xml_id {
                return Err(Error::InvalidFormat(
                    "text:changed-region text:id and xml:id must match".to_string(),
                ));
            }
            (text_id, Some(xml_id))
        },
        (Some(text_id), None) => (text_id, None),
        (None, Some(xml_id)) => (xml_id.clone(), Some(xml_id)),
        (None, None) => {
            return Err(Error::InvalidFormat(
                "text:changed-region requires text:id or xml:id".to_string(),
            ));
        },
    };
    validate_tracked_change_text(&id, "tracked-change ID", false)?;
    Ok((id, xml_id))
}

fn parse_tracked_changes_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tracked: &mut TrackedChanges,
) -> Result<()> {
    if let Some(value) = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"track-changes",
        "tracked-changes",
    )? {
        tracked.track_changes = Some(parse_tracked_change_bool("track-changes", &value)?);
    }
    if let Some(value) = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"protection-key",
        "tracked-changes",
    )? {
        validate_protection_key(&value)?;
        tracked.protection_key = Some(value);
    }
    if let Some(value) = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"protection-key-digest-algorithm",
        "tracked-changes",
    )? {
        validate_tracked_change_text(&value, "text:protection-key-digest-algorithm", false)?;
        tracked.protection_key_digest_algorithm = Some(value);
    }
    if tracked.protection_key_digest_algorithm.is_some() && tracked.protection_key.is_none() {
        return Err(Error::InvalidFormat(
            "text:protection-key-digest-algorithm requires text:protection-key".to_string(),
        ));
    }
    Ok(())
}

fn process_change_declaration_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    office_element: bool,
    dc_element: bool,
    change: &mut ActiveTrackedChange,
) -> Result<()> {
    if text_element {
        let change_type = match element.local_name().as_ref() {
            b"insertion" => Some(ChangeType::Insertion),
            b"deletion" => Some(ChangeType::Deletion),
            b"format-change" => Some(ChangeType::FormatChange),
            _ => None,
        };
        if let Some(change_type) = change_type {
            if change.change_type.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "changed region '{}' has multiple change declarations",
                    change.id
                )));
            }
            if change.depth != 2 {
                return Err(Error::InvalidFormat(
                    "change declaration must be a direct child of text:changed-region".to_string(),
                ));
            }
            match change_type {
                ChangeType::FormatChange => {
                    change.style_name = namespaced_attribute(
                        reader,
                        element,
                        TEXT_NAMESPACE,
                        b"style-name",
                        "format-change",
                    )?;
                    if let Some(style_name) = &change.style_name {
                        validate_tracked_change_text(style_name, "text:style-name", false)?;
                    }
                },
                ChangeType::Deletion => {
                    change.merge_last_paragraph = namespaced_attribute(
                        reader,
                        element,
                        TEXT_NAMESPACE,
                        b"merge-last-paragraph",
                        "deletion",
                    )?
                    .map(|value| parse_tracked_change_bool("merge-last-paragraph", &value))
                    .transpose()?;
                },
                ChangeType::Insertion => {},
            }
            change.change_type = Some(change_type);
            change.kind_depth = Some(change.depth);
            return Ok(());
        }
    }
    if office_element && element.local_name().as_ref() == b"change-info" {
        if change.kind_depth != change.depth.checked_sub(1) || change.change_info_seen {
            return Err(Error::InvalidFormat(format!(
                "invalid office:change-info in changed region '{}'",
                change.id
            )));
        }
        change.change_info_seen = true;
        change.change_info_depth = Some(change.depth);
    } else if dc_element
        && element.local_name().as_ref() == b"creator"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) || change.author.is_some() {
            return Err(Error::InvalidFormat(format!(
                "invalid dc:creator in changed region '{}'",
                change.id
            )));
        }
        change.author = Some(String::new());
        change.creator_depth = Some(change.depth);
    } else if dc_element
        && element.local_name().as_ref() == b"date"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) || change.date.is_some() {
            return Err(Error::InvalidFormat(format!(
                "invalid dc:date in changed region '{}'",
                change.id
            )));
        }
        change.date = Some(String::new());
        change.date_depth = Some(change.depth);
    } else if text_element
        && element.local_name().as_ref() == b"p"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) {
            return Err(Error::InvalidFormat(
                "change-info comment must be a direct text:p child".to_string(),
            ));
        }
        if change.comment_seen {
            append_checked(&mut change.comment, "\n")?;
        }
        change.comment_seen = true;
        change.comment_depth = Some(change.depth);
    } else if text_element
        && matches!(element.local_name().as_ref(), b"p" | b"h")
        && change.change_type == Some(ChangeType::Deletion)
        && change.change_info_depth.is_none()
        && change.paragraph_depth.is_none()
    {
        if change.seen_paragraph {
            append_checked(&mut change.content, "\n")?;
        }
        change.seen_paragraph = true;
        change.paragraph_depth = Some(change.depth);
    }
    if text_element && change.comment_depth.is_some() {
        append_text_control(reader, element, &mut change.comment)?;
    } else if text_element && change.paragraph_depth.is_some() {
        append_text_control(reader, element, &mut change.content)?;
    }
    Ok(())
}

fn process_change_declaration_empty(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    office_element: bool,
    dc_element: bool,
    change: &mut ActiveTrackedChange,
) -> Result<()> {
    if text_element
        && matches!(
            element.local_name().as_ref(),
            b"insertion" | b"deletion" | b"format-change"
        )
    {
        return Err(Error::InvalidFormat(
            "change declaration requires office:change-info".to_string(),
        ));
    }
    if office_element && element.local_name().as_ref() == b"change-info" {
        if change.kind_depth != change.depth.checked_sub(1) || change.change_info_seen {
            return Err(Error::InvalidFormat(format!(
                "invalid office:change-info in changed region '{}'",
                change.id
            )));
        }
        change.change_info_seen = true;
    } else if dc_element
        && element.local_name().as_ref() == b"creator"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) || change.author.is_some() {
            return Err(Error::InvalidFormat(format!(
                "invalid dc:creator in changed region '{}'",
                change.id
            )));
        }
        change.author = Some(String::new());
    } else if dc_element
        && element.local_name().as_ref() == b"date"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) || change.date.is_some() {
            return Err(Error::InvalidFormat(format!(
                "invalid dc:date in changed region '{}'",
                change.id
            )));
        }
        change.date = Some(String::new());
    } else if text_element
        && element.local_name().as_ref() == b"p"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) {
            return Err(Error::InvalidFormat(
                "change-info comment must be a direct text:p child".to_string(),
            ));
        }
        if change.comment_seen {
            append_checked(&mut change.comment, "\n")?;
        }
        change.comment_seen = true;
    } else if text_element
        && matches!(element.local_name().as_ref(), b"p" | b"h")
        && change.change_type == Some(ChangeType::Deletion)
        && change.change_info_depth.is_none()
    {
        if change.seen_paragraph {
            append_checked(&mut change.content, "\n")?;
        }
        change.seen_paragraph = true;
    } else if text_element && change.comment_depth.is_some() {
        append_text_control(reader, element, &mut change.comment)?;
    } else if text_element && change.paragraph_depth.is_some() {
        append_text_control(reader, element, &mut change.content)?;
    }
    Ok(())
}

fn append_change_declaration_text(change: &mut ActiveTrackedChange, value: &str) -> Result<()> {
    if change.creator_depth.is_some() {
        append_checked(
            change.author.as_mut().ok_or_else(|| {
                Error::InvalidFormat("missing tracked-change creator".to_string())
            })?,
            value,
        )
    } else if change.date_depth.is_some() {
        append_checked(
            change
                .date
                .as_mut()
                .ok_or_else(|| Error::InvalidFormat("missing tracked-change date".to_string()))?,
            value,
        )
    } else if change.comment_depth.is_some() {
        append_checked(&mut change.comment, value)
    } else if change.paragraph_depth.is_some() {
        append_checked(&mut change.content, value)
    } else {
        Ok(())
    }
}

pub(crate) fn correlate_change_ranges(content: &str, changes: &mut [TrackChange]) -> Result<()> {
    let change_types: HashMap<String, ChangeType> = changes
        .iter()
        .map(|change| (change.id.clone(), change.change_type))
        .collect();
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut tracked_depth = 0usize;
    let mut paragraph_depth: Option<usize> = None;
    let mut annotation_depth = 0usize;
    let mut ranges = ChangeRangeState::default();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid change-range XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "change range")?;
                if tracked_depth > 0 {
                    tracked_depth = checked_semantic_depth(tracked_depth, "change range")?;
                } else if text_element && element.local_name().as_ref() == b"tracked-changes" {
                    tracked_depth = 1;
                } else if annotation_depth > 0 {
                    annotation_depth = checked_semantic_depth(annotation_depth, "annotation")?;
                    if let Some(depth) = paragraph_depth.as_mut() {
                        *depth = checked_semantic_depth(*depth, "change-range paragraph")?;
                    }
                } else {
                    enter_change_paragraph(
                        text_element,
                        element,
                        &mut paragraph_depth,
                        &mut ranges.pending,
                    )?;
                    if office_element && element.local_name().as_ref() == b"annotation" {
                        annotation_depth = 1;
                    } else {
                        process_change_marker(
                            &reader,
                            element,
                            text_element,
                            paragraph_depth.is_some(),
                            &change_types,
                            &mut ranges,
                        )?;
                        if text_element && paragraph_depth.is_some() {
                            for range in ranges.pending.values_mut() {
                                append_text_control(&reader, element, &mut range.text)?;
                            }
                        }
                    }
                }
            },
            Event::Empty(ref element) if tracked_depth == 0 => {
                if annotation_depth > 0
                    || office_element && element.local_name().as_ref() == b"annotation"
                {
                    // Annotation definitions are metadata rather than visible range text.
                } else if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                    for range in ranges.pending.values_mut() {
                        if range.seen_paragraph {
                            append_checked(&mut range.text, "\n")?;
                        }
                        range.seen_paragraph = true;
                    }
                } else {
                    process_change_marker(
                        &reader,
                        element,
                        text_element,
                        paragraph_depth.is_some(),
                        &change_types,
                        &mut ranges,
                    )?;
                    if text_element && paragraph_depth.is_some() {
                        for range in ranges.pending.values_mut() {
                            append_text_control(&reader, element, &mut range.text)?;
                        }
                    }
                }
            },
            Event::Text(ref value)
                if tracked_depth == 0 && annotation_depth == 0 && paragraph_depth.is_some() =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid change-range text: {error}"))
                    })?;
                for range in ranges.pending.values_mut() {
                    append_checked(&mut range.text, &value)?;
                }
            },
            Event::CData(ref value)
                if tracked_depth == 0 && annotation_depth == 0 && paragraph_depth.is_some() =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid change-range CDATA: {error}"))
                    })?;
                for range in ranges.pending.values_mut() {
                    append_checked(&mut range.text, &value)?;
                }
            },
            Event::GeneralRef(ref reference)
                if tracked_depth == 0 && annotation_depth == 0 && paragraph_depth.is_some() =>
            {
                let value = decode_reference(reference, "change range")?;
                for range in ranges.pending.values_mut() {
                    append_checked(&mut range.text, &value)?;
                }
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("change-range XML stack underflow".to_string())
                })?;
                if tracked_depth > 0 {
                    tracked_depth = tracked_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("tracked-change range stack underflow".to_string())
                    })?;
                } else {
                    annotation_depth = annotation_depth.saturating_sub(1);
                    if let Some(depth) = paragraph_depth.as_mut() {
                        *depth = depth.checked_sub(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "change-range paragraph stack underflow".to_string(),
                            )
                        })?;
                        if *depth == 0 {
                            paragraph_depth = None;
                        }
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0
        || tracked_depth != 0
        || paragraph_depth.is_some()
        || annotation_depth != 0
    {
        return Err(Error::InvalidFormat(
            "incomplete change-range XML structure".to_string(),
        ));
    }
    if let Some(id) = ranges.pending.keys().next() {
        return Err(Error::InvalidFormat(format!(
            "unclosed text:change-start for '{id}'"
        )));
    }
    for change in changes {
        if change.change_type != ChangeType::Deletion
            && let Some(completed) = ranges.completed.remove(&change.id)
        {
            change.content.clear();
            for (index, range) in completed.iter().enumerate() {
                if index > 0 {
                    append_checked(&mut change.content, "\n")?;
                }
                append_checked(&mut change.content, range)?;
            }
        }
    }
    Ok(())
}

fn enter_change_paragraph(
    text_element: bool,
    element: &BytesStart<'_>,
    paragraph_depth: &mut Option<usize>,
    pending: &mut HashMap<String, PendingChangeRange>,
) -> Result<()> {
    if let Some(depth) = paragraph_depth.as_mut() {
        *depth = checked_semantic_depth(*depth, "change-range paragraph")?;
    } else if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
        *paragraph_depth = Some(1);
        for range in pending.values_mut() {
            if range.seen_paragraph {
                append_checked(&mut range.text, "\n")?;
            }
            range.seen_paragraph = true;
        }
    }
    Ok(())
}

fn process_change_marker(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    in_paragraph: bool,
    change_types: &HashMap<String, ChangeType>,
    ranges: &mut ChangeRangeState,
) -> Result<()> {
    if !text_element {
        return Ok(());
    }
    let marker = element.local_name();
    if !matches!(marker.as_ref(), b"change" | b"change-start" | b"change-end") {
        return Ok(());
    }
    let id = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"change-id",
        "change marker",
    )?
    .ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:{} requires text:change-id",
            String::from_utf8_lossy(marker.as_ref())
        ))
    })?;
    if !change_types.contains_key(&id) {
        return Err(Error::InvalidFormat(format!(
            "change marker references unknown ID '{id}'"
        )));
    }
    match marker.as_ref() {
        b"change-start" => {
            if ranges.pending.len() >= MAX_SEMANTIC_ITEMS {
                return Err(Error::InvalidFormat(format!(
                    "document exceeds {MAX_SEMANTIC_ITEMS} open change ranges"
                )));
            }
            if ranges
                .pending
                .insert(
                    id.clone(),
                    PendingChangeRange {
                        text: String::new(),
                        seen_paragraph: in_paragraph,
                    },
                )
                .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "duplicate open change range '{id}'"
                )));
            }
        },
        b"change-end" => {
            let range = ranges.pending.remove(&id).ok_or_else(|| {
                Error::InvalidFormat(format!("text:change-end has no open range for '{id}'"))
            })?;
            ranges.completed_count = ranges.completed_count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("completed change-range count overflow".to_string())
            })?;
            if ranges.completed_count > MAX_SEMANTIC_ITEMS {
                return Err(Error::InvalidFormat(format!(
                    "document exceeds {MAX_SEMANTIC_ITEMS} completed change ranges"
                )));
            }
            ranges.completed.entry(id).or_default().push(range.text);
        },
        b"change" => {},
        _ => unreachable!(),
    }
    Ok(())
}

pub(crate) fn parse_comments(content: &str) -> Result<Vec<Comment>> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active: Option<ActiveComment> = None;
    let mut comments = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid annotation XML: {error}")))?;
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let dc_element = is_bound(&namespace, DC_NAMESPACE);
        let meta_element = is_bound(&namespace, META_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "annotation")?;
                if let Some(comment) = active.as_mut() {
                    if office_element && element.local_name().as_ref() == b"annotation" {
                        return Err(Error::InvalidFormat(
                            "nested office:annotation elements are not allowed".to_string(),
                        ));
                    }
                    comment.depth += 1;
                    if dc_element && element.local_name().as_ref() == b"creator" {
                        comment.creator_depth = Some(comment.depth);
                    } else if dc_element && element.local_name().as_ref() == b"date" {
                        comment.date_depth = Some(comment.depth);
                    } else if meta_element && element.local_name().as_ref() == b"date-string" {
                        comment.fallback_date_depth = Some(comment.depth);
                    } else if text_element
                        && matches!(element.local_name().as_ref(), b"p" | b"h")
                        && comment.paragraph_depth.is_none()
                    {
                        if comment.seen_paragraph {
                            append_checked(&mut comment.comment.content, "\n")?;
                        }
                        comment.seen_paragraph = true;
                        comment.paragraph_depth = Some(comment.depth);
                    }
                    if text_element && comment.paragraph_depth.is_some() {
                        append_text_control(&reader, element, &mut comment.comment.content)?;
                    }
                } else if office_element && element.local_name().as_ref() == b"annotation" {
                    if comments.len() >= MAX_SEMANTIC_ITEMS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_SEMANTIC_ITEMS} annotations"
                        )));
                    }
                    let id = namespaced_attribute(
                        &reader,
                        element,
                        OFFICE_NAMESPACE,
                        b"name",
                        "annotation",
                    )?
                    .unwrap_or_else(|| format!("comment_{}", comments.len()));
                    active = Some(ActiveComment::new(id));
                }
            },
            Event::Empty(ref element) if active.is_some() => {
                let comment = active
                    .as_mut()
                    .ok_or_else(|| Error::InvalidFormat("missing active annotation".to_string()))?;
                if office_element && element.local_name().as_ref() == b"annotation" {
                    return Err(Error::InvalidFormat(
                        "nested office:annotation elements are not allowed".to_string(),
                    ));
                }
                if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                    if comment.seen_paragraph {
                        append_checked(&mut comment.comment.content, "\n")?;
                    }
                    comment.seen_paragraph = true;
                } else if text_element && comment.paragraph_depth.is_some() {
                    append_text_control(&reader, element, &mut comment.comment.content)?;
                }
            },
            Event::Empty(ref element)
                if office_element && element.local_name().as_ref() == b"annotation" =>
            {
                if comments.len() >= MAX_SEMANTIC_ITEMS {
                    return Err(Error::InvalidFormat(format!(
                        "document exceeds {MAX_SEMANTIC_ITEMS} annotations"
                    )));
                }
                let id = namespaced_attribute(
                    &reader,
                    element,
                    OFFICE_NAMESPACE,
                    b"name",
                    "annotation",
                )?
                .unwrap_or_else(|| format!("comment_{}", comments.len()));
                comments.push(Comment {
                    id,
                    author: None,
                    date: None,
                    content: String::new(),
                    reference: None,
                });
            },
            Event::Text(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid annotation text: {error}"))
                    })?;
                append_comment_text(
                    active.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("missing active annotation".to_string())
                    })?,
                    &value,
                )?;
            },
            Event::CData(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid annotation CDATA: {error}"))
                    })?;
                append_comment_text(
                    active.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("missing active annotation".to_string())
                    })?,
                    &value,
                )?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let value = decode_reference(reference, "annotation")?;
                append_comment_text(
                    active.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("missing active annotation".to_string())
                    })?,
                    &value,
                )?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("annotation XML stack underflow".to_string())
                })?;
                if let Some(comment) = active.as_mut() {
                    if comment.creator_depth == Some(comment.depth) {
                        comment.creator_depth = None;
                    }
                    if comment.date_depth == Some(comment.depth) {
                        comment.date_depth = None;
                    }
                    if comment.fallback_date_depth == Some(comment.depth) {
                        comment.fallback_date_depth = None;
                    }
                    if comment.paragraph_depth == Some(comment.depth) {
                        comment.paragraph_depth = None;
                    }
                    comment.depth = comment.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("annotation element stack underflow".to_string())
                    })?;
                    if comment.depth == 0 {
                        comments.push(
                            active
                                .take()
                                .ok_or_else(|| {
                                    Error::InvalidFormat("missing completed annotation".to_string())
                                })?
                                .finish(),
                        );
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if document_depth != 0 || active.is_some() {
        return Err(Error::InvalidFormat(
            "incomplete annotation XML structure".to_string(),
        ));
    }
    let references = parse_annotation_references(content)?;
    for comment in &mut comments {
        if let Some(reference) = references.get(&comment.id) {
            comment.reference = Some(reference.clone());
        }
    }
    Ok(comments)
}

fn append_comment_text(comment: &mut ActiveComment, value: &str) -> Result<()> {
    if comment.creator_depth.is_some() {
        let author = comment.comment.author.get_or_insert_with(String::new);
        append_checked(author, value)
    } else if comment.date_depth.is_some() {
        let date = comment.comment.date.get_or_insert_with(String::new);
        append_checked(date, value)
    } else if comment.fallback_date_depth.is_some() {
        append_checked(&mut comment.fallback_date, value)
    } else if comment.paragraph_depth.is_some() {
        append_checked(&mut comment.comment.content, value)
    } else {
        Ok(())
    }
}

fn parse_annotation_references(content: &str) -> Result<HashMap<String, String>> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut paragraph_depth: Option<usize> = None;
    let mut annotation_depth = 0usize;
    let mut annotation_name = None;
    let mut pending: HashMap<String, PendingAnnotation> = HashMap::new();
    let mut references = HashMap::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid annotation range XML: {error}"))
            })?;
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "annotation range")?;
                if let Some(depth) = paragraph_depth.as_mut() {
                    *depth += 1;
                } else if text_element
                    && matches!(element.local_name().as_ref(), b"p" | b"h")
                    && annotation_depth == 0
                {
                    paragraph_depth = Some(1);
                    for item in pending.values_mut() {
                        if item.seen_paragraph {
                            append_checked(&mut item.text, "\n")?;
                        }
                        item.seen_paragraph = true;
                    }
                }
                if annotation_depth > 0 {
                    annotation_depth += 1;
                } else if office_element && element.local_name().as_ref() == b"annotation" {
                    annotation_name = namespaced_attribute(
                        &reader,
                        element,
                        OFFICE_NAMESPACE,
                        b"name",
                        "annotation",
                    )?;
                    annotation_depth = 1;
                } else if office_element && element.local_name().as_ref() == b"annotation-end" {
                    finish_annotation_reference(&reader, element, &mut pending, &mut references)?;
                } else if text_element && paragraph_depth.is_some() {
                    for item in pending.values_mut() {
                        append_text_control(&reader, element, &mut item.text)?;
                    }
                }
            },
            Event::Empty(ref element) if annotation_depth == 0 => {
                if office_element && element.local_name().as_ref() == b"annotation" {
                    if let Some(name) = namespaced_attribute(
                        &reader,
                        element,
                        OFFICE_NAMESPACE,
                        b"name",
                        "annotation",
                    )? {
                        add_pending_annotation(
                            &mut pending,
                            PendingAnnotation::new(name, paragraph_depth.is_some()),
                        )?;
                    }
                } else if office_element && element.local_name().as_ref() == b"annotation-end" {
                    finish_annotation_reference(&reader, element, &mut pending, &mut references)?;
                } else if text_element && paragraph_depth.is_some() {
                    for item in pending.values_mut() {
                        append_text_control(&reader, element, &mut item.text)?;
                    }
                }
            },
            Event::Text(ref value) if annotation_depth == 0 && paragraph_depth.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid annotation range text: {error}"))
                    })?;
                for item in pending.values_mut() {
                    append_checked(&mut item.text, &value)?;
                }
            },
            Event::CData(ref value) if annotation_depth == 0 && paragraph_depth.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid annotation range CDATA: {error}"))
                    })?;
                for item in pending.values_mut() {
                    append_checked(&mut item.text, &value)?;
                }
            },
            Event::GeneralRef(ref reference)
                if annotation_depth == 0 && paragraph_depth.is_some() =>
            {
                let value = decode_reference(reference, "annotation range")?;
                for item in pending.values_mut() {
                    append_checked(&mut item.text, &value)?;
                }
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("annotation range XML stack underflow".to_string())
                })?;
                if annotation_depth > 0 {
                    annotation_depth -= 1;
                    if annotation_depth == 0
                        && let Some(name) = annotation_name.take()
                    {
                        add_pending_annotation(
                            &mut pending,
                            PendingAnnotation::new(name, paragraph_depth.is_some()),
                        )?;
                    }
                }
                if let Some(depth) = paragraph_depth.as_mut() {
                    *depth = (*depth).checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("annotation paragraph stack underflow".to_string())
                    })?;
                    if *depth == 0 {
                        paragraph_depth = None;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || paragraph_depth.is_some() || annotation_depth != 0 {
        return Err(Error::InvalidFormat(
            "incomplete annotation range XML structure".to_string(),
        ));
    }
    Ok(references)
}

fn finish_annotation_reference(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    pending: &mut HashMap<String, PendingAnnotation>,
    references: &mut HashMap<String, String>,
) -> Result<()> {
    let name = namespaced_attribute(reader, element, OFFICE_NAMESPACE, b"name", "annotation-end")?
        .ok_or_else(|| {
            Error::InvalidFormat("office:annotation-end requires office:name".to_string())
        })?;
    if let Some(item) = pending.remove(&name) {
        references.insert(item.name, item.text);
    }
    Ok(())
}

fn add_pending_annotation(
    pending: &mut HashMap<String, PendingAnnotation>,
    annotation: PendingAnnotation,
) -> Result<()> {
    ensure_pending_capacity(pending.len())?;
    if pending.contains_key(&annotation.name) {
        return Err(Error::InvalidFormat(format!(
            "duplicate open annotation range '{}'",
            annotation.name
        )));
    }
    pending.insert(annotation.name.clone(), annotation);
    Ok(())
}

pub(crate) fn parse_sections(content: &str) -> Result<Vec<Section>> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active: Vec<ActiveSection> = Vec::new();
    let mut sections = Vec::new();
    let mut next_order = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid section XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "section")?;
                for section in &mut active {
                    if section.source_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "section source declarations must be empty".to_string(),
                        ));
                    }
                    section.depth = checked_semantic_depth(section.depth, "section")?;
                    if text_element
                        && matches!(element.local_name().as_ref(), b"p" | b"h")
                        && section.paragraph_depth.is_none()
                    {
                        if section.seen_paragraph {
                            append_checked(&mut section.section.content, "\n")?;
                        }
                        section.seen_paragraph = true;
                        section.paragraph_depth = Some(section.depth);
                    }
                    if text_element && section.paragraph_depth.is_some() {
                        append_text_control(&reader, element, &mut section.section.content)?;
                    }
                }
                if text_element && element.local_name().as_ref() == b"section" {
                    if next_order >= MAX_SEMANTIC_ITEMS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_SEMANTIC_ITEMS} sections"
                        )));
                    }
                    active.push(ActiveSection::new(
                        section_from_start(&reader, element)?,
                        next_order,
                    ));
                    next_order += 1;
                } else if let Some(section) = active.last_mut()
                    && section.depth == 2
                    && ((text_element && element.local_name().as_ref() == b"section-source")
                        || (office_element && element.local_name().as_ref() == b"dde-source"))
                {
                    apply_section_source(
                        &reader,
                        element,
                        text_element,
                        office_element,
                        &mut section.section,
                    )?;
                    section.source_depth = Some(section.depth);
                }
            },
            Event::Empty(ref element) => {
                if text_element && element.local_name().as_ref() == b"section" {
                    if next_order >= MAX_SEMANTIC_ITEMS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_SEMANTIC_ITEMS} sections"
                        )));
                    }
                    sections.push((next_order, section_from_start(&reader, element)?));
                    next_order += 1;
                } else if let Some(section) = active.last_mut()
                    && section.depth == 1
                    && ((text_element && element.local_name().as_ref() == b"section-source")
                        || (office_element && element.local_name().as_ref() == b"dde-source"))
                {
                    apply_section_source(
                        &reader,
                        element,
                        text_element,
                        office_element,
                        &mut section.section,
                    )?;
                } else {
                    for section in &mut active {
                        if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                            if section.seen_paragraph {
                                append_checked(&mut section.section.content, "\n")?;
                            }
                            section.seen_paragraph = true;
                        } else if text_element && section.paragraph_depth.is_some() {
                            append_text_control(&reader, element, &mut section.section.content)?;
                        }
                    }
                }
            },
            Event::Text(ref value) if !active.is_empty() => {
                if active.iter().any(|section| section.source_depth.is_some()) {
                    return Err(Error::InvalidFormat(
                        "section source declarations must be empty".to_string(),
                    ));
                }
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid section text: {error}"))
                    })?;
                for section in &mut active {
                    if section.paragraph_depth.is_some() {
                        append_checked(&mut section.section.content, &value)?;
                    }
                }
            },
            Event::CData(ref value) if !active.is_empty() => {
                if active.iter().any(|section| section.source_depth.is_some()) {
                    return Err(Error::InvalidFormat(
                        "section source declarations must be empty".to_string(),
                    ));
                }
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid section CDATA: {error}"))
                    })?;
                for section in &mut active {
                    if section.paragraph_depth.is_some() {
                        append_checked(&mut section.section.content, &value)?;
                    }
                }
            },
            Event::GeneralRef(ref reference) if !active.is_empty() => {
                if active.iter().any(|section| section.source_depth.is_some()) {
                    return Err(Error::InvalidFormat(
                        "section source declarations must be empty".to_string(),
                    ));
                }
                let value = decode_reference(reference, "section")?;
                for section in &mut active {
                    if section.source_depth == Some(section.depth) {
                        section.source_depth = None;
                    }
                    if section.paragraph_depth.is_some() {
                        append_checked(&mut section.section.content, &value)?;
                    }
                }
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("section XML stack underflow".to_string())
                })?;
                for section in &mut active {
                    if section.paragraph_depth == Some(section.depth) {
                        section.paragraph_depth = None;
                    }
                    section.depth = section.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("section element stack underflow".to_string())
                    })?;
                }
                if let Some(section) = active.pop_if(|section| section.depth == 0) {
                    sections.push(section.into_ordered());
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete section XML structure".to_string(),
        ));
    }
    sections.sort_by_key(|(order, _)| *order);
    Ok(sections.into_iter().map(|(_, section)| section).collect())
}

fn section_from_start(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Section> {
    let name = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"name", "section")?
        .ok_or_else(|| Error::InvalidFormat("text:section requires text:name".to_string()))?;
    let style = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"style-name", "section")?;
    let protected = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"protected", "section")?
        .map(|value| parse_boolean(&value, "text:protected"))
        .transpose()?
        .unwrap_or(false);
    let xml_id = namespaced_attribute(reader, element, XML_NAMESPACE, b"id", "section")?;
    let protection_key = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"protection-key",
        "section",
    )?;
    let protection_key_digest_algorithm = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"protection-key-digest-algorithm",
        "section",
    )?;
    let display_value =
        namespaced_attribute(reader, element, TEXT_NAMESPACE, b"display", "section")?;
    let condition = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"condition", "section")?;
    let display = match display_value.as_deref() {
        None | Some("true") => SectionDisplay::Visible,
        Some("none") => SectionDisplay::Hidden,
        Some("condition") if condition.is_some() => SectionDisplay::Condition,
        Some("condition") => {
            return Err(Error::InvalidFormat(
                "text:display='condition' requires text:condition".to_string(),
            ));
        },
        Some(value) => {
            return Err(Error::InvalidFormat(format!(
                "unsupported text:display value '{value}'"
            )));
        },
    };
    if condition.is_some() && display != SectionDisplay::Condition {
        return Err(Error::InvalidFormat(
            "text:condition requires text:display='condition'".to_string(),
        ));
    }
    Ok(Section {
        name,
        style,
        protected,
        xml_id,
        protection_key,
        protection_key_digest_algorithm,
        display,
        condition,
        source: None,
        dde_source: None,
        content: String::new(),
    })
}

fn apply_section_source(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    office_element: bool,
    section: &mut Section,
) -> Result<()> {
    if section.source.is_some() || section.dde_source.is_some() {
        return Err(Error::InvalidFormat(
            "section may have only one source declaration".to_string(),
        ));
    }
    if text_element && element.local_name().as_ref() == b"section-source" {
        let href =
            namespaced_attribute(reader, element, XLINK_NAMESPACE, b"href", "section source")?;
        let link_type =
            namespaced_attribute(reader, element, XLINK_NAMESPACE, b"type", "section source")?;
        if link_type.as_deref().is_some_and(|value| value != "simple") {
            return Err(Error::InvalidFormat(
                "section source xlink:type must be 'simple'".to_string(),
            ));
        }
        let show =
            namespaced_attribute(reader, element, XLINK_NAMESPACE, b"show", "section source")?;
        if show.as_deref().is_some_and(|value| value != "embed") {
            return Err(Error::InvalidFormat(
                "section source xlink:show must be 'embed'".to_string(),
            ));
        }
        if href.is_some() != link_type.is_some() || href.is_none() && show.is_some() {
            return Err(Error::InvalidFormat(
                "section source xlink:href and xlink:type must appear together".to_string(),
            ));
        }
        section.source = Some(SectionSource {
            href,
            section_name: namespaced_attribute(
                reader,
                element,
                TEXT_NAMESPACE,
                b"section-name",
                "section source",
            )?,
            filter_name: namespaced_attribute(
                reader,
                element,
                TEXT_NAMESPACE,
                b"filter-name",
                "section source",
            )?,
        });
    } else if office_element && element.local_name().as_ref() == b"dde-source" {
        let conversion_mode = namespaced_attribute(
            reader,
            element,
            OFFICE_NAMESPACE,
            b"conversion-mode",
            "section DDE source",
        )?;
        if conversion_mode.as_deref().is_some_and(|value| {
            !matches!(
                value,
                "into-default-style-data-style" | "into-english-number" | "keep-text"
            )
        }) {
            return Err(Error::InvalidFormat(
                "unsupported office:conversion-mode on section DDE source".to_string(),
            ));
        }
        section.dde_source = Some(SectionDdeSource {
            name: namespaced_attribute(
                reader,
                element,
                OFFICE_NAMESPACE,
                b"name",
                "section DDE source",
            )?,
            conversion_mode,
            automatic_update: namespaced_attribute(
                reader,
                element,
                OFFICE_NAMESPACE,
                b"automatic-update",
                "section DDE source",
            )?
            .map(|value| parse_boolean(&value, "office:automatic-update"))
            .transpose()?,
        });
    }
    Ok(())
}
