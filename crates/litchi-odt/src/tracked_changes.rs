//! Validation and deterministic serialization for text change declarations.

use super::parser::{ChangeType, TrackChange, TrackedChanges};
use crate::elements::xml::{
    OFFICE_NAMESPACE, TEXT_NAMESPACE, XML_NAMESPACE, is_bound, namespaced_attribute,
};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::ops::Range;

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";

const MAX_CHANGES: usize = 1_000_000;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

impl TrackedChanges {
    /// Validate a manually constructed tracked-change declaration table.
    ///
    /// This validates the declaration metadata only. Range markers remain part of
    /// the surrounding document and are correlated by [`crate::Document::tracked_changes`]
    /// when parsing a complete document.
    pub fn validate(&self) -> Result<()> {
        if self.changes.len() > MAX_CHANGES {
            return invalid(format!(
                "tracked changes exceed the {MAX_CHANGES} declaration limit"
            ));
        }
        validate_policy(self)?;

        let mut ids = HashSet::with_capacity(self.changes.len());
        let mut aggregate = 0usize;
        for change in &self.changes {
            change.validate()?;
            if !ids.insert(change.id.as_str()) {
                return invalid(format!("duplicate tracked-change ID '{}'", change.id));
            }
            aggregate = aggregate
                .checked_add(change.content.len())
                .and_then(|value| value.checked_add(change.author.as_deref().map_or(0, str::len)))
                .and_then(|value| value.checked_add(change.date.as_deref().map_or(0, str::len)))
                .and_then(|value| value.checked_add(change.comment.as_deref().map_or(0, str::len)))
                .ok_or_else(|| make_error("tracked-change aggregate size overflow"))?;
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("tracked-change values exceed 16 MiB");
            }
        }
        Ok(())
    }

    /// Serialize the `text:tracked-changes` declaration table deterministically.
    ///
    /// Insertion and format-change `content` is live document text discovered from
    /// range markers, so it is not duplicated in the declaration table. Deletion
    /// content is emitted as inert ODF text with tabs, line breaks, and repeated
    /// spaces represented by their standard text controls.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::with_capacity(256 + self.changes.len().saturating_mul(192));
        xml.push_str(
            r#"<text:tracked-changes xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/""#,
        );
        if let Some(value) = self.track_changes {
            xml.push_str(r#" text:track-changes=""#);
            xml.push_str(if value { "true" } else { "false" });
            xml.push('"');
        }
        if let Some(value) = &self.protection_key {
            push_attribute(&mut xml, "text:protection-key", value);
        }
        if let Some(value) = &self.protection_key_digest_algorithm {
            push_attribute(&mut xml, "text:protection-key-digest-algorithm", value);
        }
        if self.changes.is_empty() {
            xml.push_str("/>");
            return Ok(xml);
        }
        xml.push('>');
        for change in &self.changes {
            write_change(&mut xml, change);
        }
        xml.push_str("</text:tracked-changes>");
        Ok(xml)
    }
}

impl TrackChange {
    /// Validate one change declaration before serialization.
    pub fn validate(&self) -> Result<()> {
        validate_ncname(&self.id, "tracked-change ID")?;
        if let Some(xml_id) = &self.xml_id {
            validate_ncname(xml_id, "tracked-change xml:id")?;
            if xml_id != &self.id {
                return invalid("tracked-change text:id and xml:id must match");
            }
        }
        let author = self
            .author
            .as_deref()
            .ok_or_else(|| make_error("tracked change requires dc:creator"))?;
        let date = self
            .date
            .as_deref()
            .ok_or_else(|| make_error("tracked change requires dc:date"))?;
        validate_value(author, "tracked-change creator", false)?;
        validate_value(date, "tracked-change date", false)?;
        if let Some(comment) = &self.comment {
            validate_value(comment, "tracked-change comment", true)?;
        }
        validate_value(&self.content, "tracked-change content", true)?;

        match self.change_type {
            ChangeType::Insertion => {
                if self.style_name.is_some() || self.merge_last_paragraph.is_some() {
                    return invalid("insertion cannot carry format-change or deletion attributes");
                }
            },
            ChangeType::Deletion => {
                if self.style_name.is_some() {
                    return invalid("deletion cannot carry text:style-name");
                }
            },
            ChangeType::FormatChange => {
                if self.merge_last_paragraph.is_some() {
                    return invalid("format change cannot carry text:merge-last-paragraph");
                }
                if let Some(style_name) = &self.style_name {
                    validate_value(style_name, "format-change style name", false)?;
                }
            },
        }
        Ok(())
    }
}

fn validate_policy(value: &TrackedChanges) -> Result<()> {
    match (
        value.protection_key.as_deref(),
        value.protection_key_digest_algorithm.as_deref(),
    ) {
        (None, Some(_)) => {
            return invalid("protection-key digest algorithm requires a protection key");
        },
        (Some(key), algorithm) => {
            validate_value(key, "tracked-change protection key", false)?;
            BASE64_STANDARD.decode(key).map_err(|error| {
                make_error(format!("invalid tracked-change protection key: {error}"))
            })?;
            if let Some(algorithm) = algorithm {
                validate_value(
                    algorithm,
                    "tracked-change protection-key digest algorithm",
                    false,
                )?;
            }
        },
        (None, None) => {},
    }
    Ok(())
}

fn write_change(xml: &mut String, change: &TrackChange) {
    xml.push_str(r#"<text:changed-region text:id=""#);
    push_escaped(xml, &change.id, true);
    xml.push('"');
    if let Some(xml_id) = &change.xml_id {
        push_attribute(xml, "xml:id", xml_id);
    }
    xml.push('>');

    let kind = match change.change_type {
        ChangeType::Insertion => "insertion",
        ChangeType::Deletion => "deletion",
        ChangeType::FormatChange => "format-change",
    };
    xml.push_str("<text:");
    xml.push_str(kind);
    if let Some(style_name) = &change.style_name {
        push_attribute(xml, "text:style-name", style_name);
    }
    if let Some(value) = change.merge_last_paragraph {
        xml.push_str(r#" text:merge-last-paragraph=""#);
        xml.push_str(if value { "true" } else { "false" });
        xml.push('"');
    }
    xml.push('>');
    xml.push_str("<office:change-info><dc:creator>");
    push_escaped(
        xml,
        change.author.as_deref().expect("validated creator"),
        false,
    );
    xml.push_str("</dc:creator><dc:date>");
    push_escaped(xml, change.date.as_deref().expect("validated date"), false);
    xml.push_str("</dc:date></office:change-info>");
    if let Some(comment) = &change.comment {
        xml.truncate(xml.len() - "</office:change-info>".len());
        xml.push_str("<text:p>");
        write_text_content(xml, comment);
        xml.push_str("</text:p></office:change-info>");
    }
    if change.change_type == ChangeType::Deletion && !change.content.is_empty() {
        xml.push_str("<text:p>");
        write_text_content(xml, &change.content);
        xml.push_str("</text:p>");
    }
    xml.push_str("</text:");
    xml.push_str(kind);
    xml.push_str("></text:changed-region>");
}

fn write_text_content(xml: &mut String, value: &str) {
    let mut plain_start = 0usize;
    let mut chars = value.char_indices().peekable();
    while let Some((index, character)) = chars.next() {
        let special = matches!(character, ' ' | '\t' | '\n' | '\r');
        if !special {
            continue;
        }
        if plain_start < index {
            push_escaped(xml, &value[plain_start..index], false);
        }
        match character {
            ' ' => {
                let mut count = 1usize;
                let mut end = index + 1;
                while let Some(&(next_index, ' ')) = chars.peek() {
                    chars.next();
                    count += 1;
                    end = next_index + 1;
                }
                if count == 1 {
                    xml.push(' ');
                } else {
                    xml.push_str(r#"<text:s text:c=""#);
                    xml.push_str(&count.to_string());
                    xml.push_str("\"/>");
                }
                plain_start = end;
            },
            '\t' => {
                xml.push_str("<text:tab/>");
                plain_start = index + 1;
            },
            '\n' => {
                xml.push_str("<text:line-break/>");
                plain_start = index + 1;
            },
            '\r' => {
                if let Some(&(next_index, '\n')) = chars.peek() {
                    chars.next();
                    plain_start = next_index + 1;
                } else {
                    plain_start = index + 1;
                }
                xml.push_str("<text:line-break/>");
            },
            _ => unreachable!(),
        }
    }
    if plain_start < value.len() {
        push_escaped(xml, &value[plain_start..], false);
    }
}

fn push_attribute(xml: &mut String, name: &str, value: &str) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    push_escaped(xml, value, true);
    xml.push('"');
}

fn push_escaped(xml: &mut String, value: &str, attribute: bool) {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' if attribute => xml.push_str("&quot;"),
            '\'' if attribute => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
}

fn validate_ncname(value: &str, context: &str) -> Result<()> {
    validate_value(value, context, false)?;
    let mut characters = value.chars();
    let first = characters.next().expect("non-empty value validated");
    if !(first == '_' || first.is_alphabetic())
        || characters.any(|character| {
            !(character == '_'
                || character == '-'
                || character == '.'
                || character.is_alphanumeric())
        })
    {
        return invalid(format!("{context} must be an XML NCName"));
    }
    Ok(())
}

fn validate_value(value: &str, context: &str, empty_allowed: bool) -> Result<()> {
    if !empty_allowed && value.is_empty() {
        return invalid(format!("{context} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{context} exceeds 64 KiB"));
    }
    if value.chars().any(
        |character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'),
    ) {
        return invalid(format!("{context} contains an XML-prohibited character"));
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Stable text story used when placing tracked-change markers.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum OdtTrackedStory {
    /// Body paragraph or heading outside tables, in document order.
    Paragraph(usize),
    /// Paragraph inside a table cell, addressed by lexical table/row/cell/paragraph order.
    TableCell {
        table: usize,
        row: usize,
        cell: usize,
        paragraph: usize,
    },
}

/// Unicode-scalar position inside one stable text story.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdtTrackedPosition {
    pub story: OdtTrackedStory,
    pub character: usize,
}

#[derive(Clone)]
struct StorySite {
    story: OdtTrackedStory,
    boundaries: Vec<Option<usize>>,
    empty: Option<(Range<usize>, String)>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum MarkerKind {
    Point,
    Start,
    End,
}

struct MarkerSite {
    id: String,
    kind: MarkerKind,
    span: Range<usize>,
}

struct XmlSites {
    office_text_open_end: usize,
    tracked_changes: Option<Range<usize>>,
    markers: Vec<MarkerSite>,
    stories: Vec<StorySite>,
}

struct ActiveStory {
    story: OdtTrackedStory,
    depth: usize,
    boundaries: Vec<Option<usize>>,
}

struct TableContext {
    table: usize,
    next_row: usize,
    row: Option<usize>,
    next_cell: usize,
    cell: Option<usize>,
    next_paragraph: usize,
}

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
    start: &OdtTrackedPosition,
    end: &OdtTrackedPosition,
) -> Result<String> {
    let tracked = super::parser::OdtParser::parse_tracked_changes(xml)?;
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
    position: &OdtTrackedPosition,
) -> Result<String> {
    let tracked = super::parser::OdtParser::parse_tracked_changes(xml)?;
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

fn validate_authored_tracked_xml(xml: &str) -> Result<TrackedChanges> {
    let tracked = super::parser::OdtParser::parse_tracked_changes(xml)?;
    tracked.validate()?;
    let types = tracked
        .changes
        .iter()
        .map(|change| (change.id.as_str(), change.change_type))
        .collect::<std::collections::HashMap<_, _>>();
    let sites = scan_mutable_tracked_xml(xml)?;
    let mut stack = Vec::<&str>::new();
    for marker in &sites.markers {
        let kind = types.get(marker.id.as_str()).ok_or_else(|| {
            make_error(format!(
                "change marker references unknown ID '{}'",
                marker.id
            ))
        })?;
        match marker.kind {
            MarkerKind::Point => {
                if *kind != ChangeType::Deletion {
                    return invalid("text:change marker requires a deletion declaration");
                }
            },
            MarkerKind::Start => {
                if *kind == ChangeType::Deletion {
                    return invalid("deletion declaration cannot use range markers");
                }
                if stack.contains(&marker.id.as_str()) {
                    return invalid("duplicate open tracked-change range");
                }
                stack.push(marker.id.as_str());
            },
            MarkerKind::End => {
                if stack.pop() != Some(marker.id.as_str()) {
                    return invalid("tracked-change ranges must be balanced and noncrossing");
                }
            },
        }
    }
    if !stack.is_empty() {
        return invalid("unclosed tracked-change range");
    }
    Ok(tracked)
}

fn resolve_story_position(stories: &[StorySite], position: &OdtTrackedPosition) -> Result<usize> {
    let site = stories
        .iter()
        .find(|site| site.story == position.story)
        .ok_or_else(|| make_error("tracked-change story was not found"))?;
    if site.empty.is_some() {
        return invalid("range markers cannot split an empty story");
    }
    site.boundaries
        .get(position.character)
        .copied()
        .flatten()
        .ok_or_else(|| {
            make_error("tracked-change offset is out of bounds or would split an XML text control")
        })
}

fn escaped_tracked_id(id: &str) -> String {
    let mut output = String::new();
    push_escaped(&mut output, id, true);
    output
}

fn scan_mutable_tracked_xml(xml: &str) -> Result<XmlSites> {
    if xml.len() > 256 * 1024 * 1024 {
        return invalid("mutable tracked-change XML exceeds 256 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous_end = 0usize;
    let mut depth = 0usize;
    let mut office_text_depth = None;
    let mut office_text_open_end = None;
    let mut tracked: Option<(usize, usize)> = None;
    let mut tracked_changes = None;
    let mut annotation_depth = None;
    let mut stories = Vec::new();
    let mut active_story: Option<ActiveStory> = None;
    let mut body_paragraph = 0usize;
    let mut next_table = 0usize;
    let mut table: Option<TableContext> = None;
    let mut markers = Vec::new();
    let mut xml_ids = HashSet::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid mutable tracked-change XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        let table_element = is_bound(&namespace, TABLE_NAMESPACE);
        drop(namespace);
        let event_end = reader.buffer_position() as usize;
        let span = previous_end..event_end;
        match event {
            Event::Start(ref element) => {
                validate_mutable_xml_id(&reader, element, &mut xml_ids)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| make_error("XML depth overflow"))?;
                if depth > 4096 {
                    return invalid("mutable tracked-change XML nesting exceeds 4096");
                }
                let local = element.local_name();
                if office_element && local.as_ref() == b"text" && office_text_depth.is_none() {
                    office_text_depth = Some(depth);
                    office_text_open_end = Some(span.end);
                }
                if text_element && local.as_ref() == b"tracked-changes" {
                    tracked = Some((depth, span.start));
                } else if office_element && local.as_ref() == b"annotation" {
                    annotation_depth = Some(depth);
                }
                if tracked.is_none() && annotation_depth.is_none() {
                    update_table_start(table_element, local.as_ref(), &mut table, &mut next_table);
                    if text_element && matches!(local.as_ref(), b"p" | b"h") {
                        let story = next_story(&mut table, &mut body_paragraph)?;
                        active_story = Some(ActiveStory {
                            story,
                            depth,
                            boundaries: vec![Some(span.end)],
                        });
                    }
                    if text_element {
                        scan_marker(&reader, element, local.as_ref(), span.clone(), &mut markers)?;
                    }
                }
            },
            Event::Empty(ref element) => {
                validate_mutable_xml_id(&reader, element, &mut xml_ids)?;
                let local = element.local_name();
                if text_element && local.as_ref() == b"tracked-changes" {
                    if tracked_changes.replace(span.clone()).is_some() {
                        return invalid("multiple text:tracked-changes elements are not allowed");
                    }
                } else if tracked.is_none() && annotation_depth.is_none() {
                    if text_element && matches!(local.as_ref(), b"p" | b"h") {
                        let story = next_story(&mut table, &mut body_paragraph)?;
                        let qname = std::str::from_utf8(element.name().as_ref())
                            .map_err(|_| make_error("non-UTF-8 story element name"))?
                            .to_string();
                        stories.push(StorySite {
                            story,
                            boundaries: vec![None],
                            empty: Some((span.clone(), qname)),
                        });
                    }
                    if text_element {
                        scan_marker(&reader, element, local.as_ref(), span.clone(), &mut markers)?;
                        append_text_control_boundaries(
                            &reader,
                            element,
                            local.as_ref(),
                            span.end,
                            active_story.as_mut(),
                        )?;
                    }
                }
            },
            Event::Text(_) | Event::GeneralRef(_) => {
                if let Some(story) = active_story.as_mut()
                    && tracked.is_none()
                    && annotation_depth.is_none()
                {
                    append_raw_text_boundaries(xml, span.clone(), &mut story.boundaries)?;
                }
            },
            Event::CData(_) => {
                if let Some(story) = active_story.as_mut()
                    && tracked.is_none()
                    && annotation_depth.is_none()
                {
                    let raw = xml
                        .get(span.clone())
                        .and_then(|value| value.strip_prefix("<![CDATA["))
                        .and_then(|value| value.strip_suffix("]]>"))
                        .ok_or_else(|| make_error("invalid CDATA event span"))?;
                    let count = raw.chars().count();
                    for _ in 1..count {
                        story.boundaries.push(None);
                    }
                    if count > 0 {
                        story.boundaries.push(Some(span.end));
                    }
                }
            },
            Event::End(ref element) => {
                let local = element.local_name();
                if active_story
                    .as_ref()
                    .is_some_and(|story| story.depth == depth)
                {
                    let mut story = active_story.take().expect("checked active story");
                    if let Some(last) = story.boundaries.last_mut() {
                        *last = Some(span.start);
                    }
                    stories.push(StorySite {
                        story: story.story,
                        boundaries: story.boundaries,
                        empty: None,
                    });
                }
                if tracked == Some((depth, tracked.map(|value| value.1).unwrap_or(0))) {
                    let (_, start) = tracked.take().expect("checked tracked container");
                    if tracked_changes.replace(start..span.end).is_some() {
                        return invalid("multiple text:tracked-changes elements are not allowed");
                    }
                }
                if annotation_depth == Some(depth) {
                    annotation_depth = None;
                }
                update_table_end(table_element, local.as_ref(), &mut table);
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| make_error("XML stack underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTD and processing instructions are not allowed");
            },
            Event::Eof => break,
            _ => {},
        }
        previous_end = event_end;
        buffer.clear();
    }
    let office_text_open_end =
        office_text_open_end.ok_or_else(|| make_error("document has no office:text body"))?;
    Ok(XmlSites {
        office_text_open_end,
        tracked_changes,
        markers,
        stories,
    })
}

fn validate_mutable_xml_id(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    ids: &mut HashSet<String>,
) -> Result<()> {
    if let Some(id) = namespaced_attribute(reader, element, XML_NAMESPACE, b"id", "xml:id")? {
        validate_ncname(&id, "xml:id")?;
        if !ids.insert(id.clone()) {
            return invalid(format!("duplicate xml:id '{id}'"));
        }
    }
    Ok(())
}

fn next_story(table: &mut Option<TableContext>, body: &mut usize) -> Result<OdtTrackedStory> {
    if let Some(table) = table {
        let row = table
            .row
            .ok_or_else(|| make_error("table paragraph is outside a row"))?;
        let cell = table
            .cell
            .ok_or_else(|| make_error("table paragraph is outside a cell"))?;
        let paragraph = table.next_paragraph;
        table.next_paragraph = table.next_paragraph.saturating_add(1);
        Ok(OdtTrackedStory::TableCell {
            table: table.table,
            row,
            cell,
            paragraph,
        })
    } else {
        let index = *body;
        *body = body.saturating_add(1);
        Ok(OdtTrackedStory::Paragraph(index))
    }
}

fn update_table_start(
    table_element: bool,
    local: &[u8],
    table: &mut Option<TableContext>,
    next_table: &mut usize,
) {
    if !table_element {
        return;
    }
    match local {
        b"table" if table.is_none() => {
            *table = Some(TableContext {
                table: *next_table,
                next_row: 0,
                row: None,
                next_cell: 0,
                cell: None,
                next_paragraph: 0,
            });
            *next_table = next_table.saturating_add(1);
        },
        b"table-row" => {
            if let Some(value) = table {
                value.row = Some(value.next_row);
                value.next_row = value.next_row.saturating_add(1);
                value.next_cell = 0;
            }
        },
        b"table-cell" | b"covered-table-cell" => {
            if let Some(value) = table {
                value.cell = Some(value.next_cell);
                value.next_cell = value.next_cell.saturating_add(1);
                value.next_paragraph = 0;
            }
        },
        _ => {},
    }
}

fn update_table_end(table_element: bool, local: &[u8], table: &mut Option<TableContext>) {
    if !table_element {
        return;
    }
    match local {
        b"table" => *table = None,
        b"table-row" => {
            if let Some(value) = table {
                value.row = None;
            }
        },
        b"table-cell" | b"covered-table-cell" => {
            if let Some(value) = table {
                value.cell = None;
            }
        },
        _ => {},
    }
}

fn scan_marker(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
    span: Range<usize>,
    markers: &mut Vec<MarkerSite>,
) -> Result<()> {
    let kind = match local {
        b"change" => MarkerKind::Point,
        b"change-start" => MarkerKind::Start,
        b"change-end" => MarkerKind::End,
        _ => return Ok(()),
    };
    let id = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"change-id",
        "change marker",
    )?
    .ok_or_else(|| make_error("change marker requires text:change-id"))?;
    markers.push(MarkerSite { id, kind, span });
    Ok(())
}

fn append_raw_text_boundaries(
    xml: &str,
    span: Range<usize>,
    boundaries: &mut Vec<Option<usize>>,
) -> Result<()> {
    let raw = xml
        .get(span.clone())
        .ok_or_else(|| make_error("invalid text event span"))?;
    let mut index = 0usize;
    while index < raw.len() {
        if raw.as_bytes()[index] == b'&' {
            let end = raw[index..]
                .find(';')
                .map(|value| index + value + 1)
                .ok_or_else(|| make_error("unterminated XML reference"))?;
            index = end;
        } else {
            let character = raw[index..]
                .chars()
                .next()
                .ok_or_else(|| make_error("invalid UTF-8 text boundary"))?;
            index += character.len_utf8();
        }
        boundaries.push(Some(span.start + index));
    }
    Ok(())
}

fn append_text_control_boundaries(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
    event_end: usize,
    story: Option<&mut ActiveStory>,
) -> Result<()> {
    let Some(story) = story else {
        return Ok(());
    };
    let count = match local {
        b"tab" | b"line-break" => 1,
        b"s" => namespaced_attribute(reader, element, TEXT_NAMESPACE, b"c", "text:s")?
            .map(|value| {
                value
                    .parse::<usize>()
                    .map_err(|_| make_error("invalid text:s count"))
            })
            .transpose()?
            .unwrap_or(1),
        _ => return Ok(()),
    };
    for _ in 1..count {
        story.boundaries.push(None);
    }
    if count > 0 {
        story.boundaries.push(Some(event_end));
    }
    Ok(())
}

fn apply_tracked_edits(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> Result<String> {
    edits.sort_by(|left, right| {
        right
            .0
            .start
            .cmp(&left.0.start)
            .then_with(|| right.0.end.cmp(&left.0.end))
    });
    let mut output = xml.to_string();
    let mut previous = xml.len();
    for (span, replacement) in edits {
        if span.start > span.end || span.end > previous || span.end > output.len() {
            return invalid("overlapping or invalid tracked-change mutation spans");
        }
        output.replace_range(span.clone(), &replacement);
        previous = span.start;
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::OdtParser;

    fn change(id: &str, change_type: ChangeType, content: &str) -> TrackChange {
        TrackChange {
            id: id.to_string(),
            xml_id: Some(id.to_string()),
            author: Some("A & B".to_string()),
            date: Some("2026-07-17T12:00:00+08:00".to_string()),
            comment: None,
            change_type,
            style_name: (change_type == ChangeType::FormatChange)
                .then(|| "Changed Style".to_string()),
            merge_last_paragraph: (change_type == ChangeType::Deletion).then_some(false),
            content: content.to_string(),
        }
    }

    #[test]
    fn serializes_and_reparses_all_declaration_kinds() {
        let declarations = TrackedChanges {
            track_changes: Some(true),
            protection_key: Some("YWJj".to_string()),
            protection_key_digest_algorithm: Some("urn:example:sha256".to_string()),
            changes: vec![
                change("insert_1", ChangeType::Insertion, "live text"),
                change("delete_1", ChangeType::Deletion, "gone  text\t<&\nnext"),
                change("format_1", ChangeType::FormatChange, "formatted"),
            ],
        };
        let xml = declarations.to_xml_fragment().unwrap();
        assert!(xml.contains("<text:s text:c=\"2\"/>"));
        assert!(xml.contains("<text:tab/>"));
        assert!(xml.contains("&lt;&amp;"));
        assert!(!xml.contains("live text"));

        let parsed = OdtParser::parse_tracked_changes(&xml).unwrap();
        assert_eq!(parsed.track_changes, Some(true));
        assert_eq!(parsed.protection_key.as_deref(), Some("YWJj"));
        assert_eq!(parsed.changes.len(), 3);
        assert_eq!(parsed.changes[1].content, "gone  text\t<&\nnext");
        assert_eq!(
            parsed.changes[2].style_name.as_deref(),
            Some("Changed Style")
        );
    }

    #[test]
    fn rejects_invalid_constructed_declarations() {
        let mut declarations = TrackedChanges {
            changes: vec![change("same", ChangeType::Insertion, "")],
            ..TrackedChanges::default()
        };
        declarations.changes.push(declarations.changes[0].clone());
        assert!(declarations.to_xml_fragment().is_err());

        declarations.changes.truncate(1);
        declarations.changes[0].id = "bad:id".to_string();
        assert!(declarations.to_xml_fragment().is_err());

        declarations.changes[0].id = "valid".to_string();
        declarations.changes[0].xml_id = Some("other".to_string());
        assert!(declarations.to_xml_fragment().is_err());

        declarations.changes[0].xml_id = Some("valid".to_string());
        declarations.protection_key_digest_algorithm = Some("urn:sha256".to_string());
        assert!(declarations.to_xml_fragment().is_err());
    }

    #[test]
    fn parses_a_libreoffice_flat_document_before_serializing() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/libreoffice-core/sw/qa/uitest/data/redline-autocorrect.fodt");
        let Ok(xml) = std::fs::read_to_string(path) else {
            return;
        };
        let parsed = OdtParser::parse_tracked_changes(&xml).unwrap();
        assert!(!parsed.changes.is_empty());
        let serialized = parsed.to_xml_fragment().unwrap();
        let reparsed = OdtParser::parse_tracked_changes(&serialized).unwrap();
        assert_eq!(reparsed.changes.len(), parsed.changes.len());
    }
}
