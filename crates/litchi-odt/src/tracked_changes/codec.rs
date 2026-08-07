//! XML codecs and lossless document mutations for tracked changes.

use super::model::{Position, Story, validate_ncname};
use super::{TABLE_NAMESPACE, invalid, make_error};
use crate::elements::xml::{
    OFFICE_NAMESPACE, TEXT_NAMESPACE, XML_NAMESPACE, is_bound, namespaced_attribute,
};
use crate::parser::{ChangeType, Parser, TrackChange, TrackedChanges};
use litchi_core::Result;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::ops::Range;
impl TrackedChanges {
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

#[derive(Clone)]
pub(super) struct StorySite {
    pub(super) story: Story,
    pub(super) boundaries: Vec<Option<usize>>,
    pub(super) empty: Option<(Range<usize>, String)>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum MarkerKind {
    Point,
    Start,
    End,
}

pub(super) struct MarkerSite {
    pub(super) id: String,
    kind: MarkerKind,
    pub(super) span: Range<usize>,
}

pub(super) struct XmlSites {
    pub(super) office_text_open_end: usize,
    pub(super) tracked_changes: Option<Range<usize>>,
    pub(super) markers: Vec<MarkerSite>,
    pub(super) stories: Vec<StorySite>,
}

struct ActiveStory {
    story: Story,
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

pub(super) fn validate_authored_tracked_xml(xml: &str) -> Result<TrackedChanges> {
    let tracked = Parser::parse_tracked_changes(xml)?;
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

pub(super) fn resolve_story_position(stories: &[StorySite], position: &Position) -> Result<usize> {
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

pub(super) fn escaped_tracked_id(id: &str) -> String {
    let mut output = String::new();
    push_escaped(&mut output, id, true);
    output
}

pub(super) fn scan_mutable_tracked_xml(xml: &str) -> Result<XmlSites> {
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
                if tracked == Some((depth, tracked.map_or(0, |value| value.1))) {
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

fn next_story(table: &mut Option<TableContext>, body: &mut usize) -> Result<Story> {
    if let Some(table) = table {
        let row = table
            .row
            .ok_or_else(|| make_error("table paragraph is outside a row"))?;
        let cell = table
            .cell
            .ok_or_else(|| make_error("table paragraph is outside a cell"))?;
        let paragraph = table.next_paragraph;
        table.next_paragraph = table.next_paragraph.saturating_add(1);
        Ok(Story::TableCell {
            table: table.table,
            row,
            cell,
            paragraph,
        })
    } else {
        let index = *body;
        *body = body.saturating_add(1);
        Ok(Story::Paragraph(index))
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

pub(super) fn apply_tracked_edits(
    xml: &str,
    mut edits: Vec<(Range<usize>, String)>,
) -> Result<String> {
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
