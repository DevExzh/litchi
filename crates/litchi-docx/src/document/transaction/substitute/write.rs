//! Planning a substitution: the splices of every changed paragraph, the
//! text each must read back as, and the runs split out for formatting.

use std::ops::Range;

use super::super::{Error, MAX_REPLACEMENT_TEXT_BYTES, Refusal, Result};
use super::scan::qualified;
use super::{Field, Outcome, Paragraph, Replacement, RunSlot, Story, TextNode};
use crate::properties;

/// One changed paragraph: its splices (absolute story offsets, ascending)
/// and the text it must read back as.
pub(super) struct Change {
    pub(super) paragraph: usize,
    pub(super) splices: Vec<(usize, usize, Vec<u8>)>,
    pub(super) text: String,
}

/// The changes of one plan and its counts.
pub(super) struct Plan {
    pub(super) changes: Vec<Change>,
    pub(super) outcome: Outcome,
}

fn refused(paragraph: usize, reason: Refusal) -> Error {
    Error::Refused {
        position: paragraph,
        reason,
    }
}

fn spliced(error: properties::SpliceError) -> Error {
    match error {
        properties::SpliceError::Refused(reason) => refused(0, reason),
        properties::SpliceError::Document(error) => Error::Document(error),
        properties::SpliceError::Transaction(error) => error,
    }
}

fn validate_text(paragraph: usize, text: &str) -> Result<()> {
    if text.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{A}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
            || character == super::OBJECT
    }) {
        return Err(refused(paragraph, Refusal::StructuralText));
    }
    Ok(())
}

/// Plan `replacements` (sorted by paragraph, then range) over `story`.
pub(super) fn plan<F: FnMut(&mut properties::Run)>(
    story: &Story,
    replacements: &[Replacement],
    format: &mut F,
) -> Result<Plan> {
    let mut changes = Vec::new();
    let mut outcome = Outcome::default();
    let mut total = 0usize;
    let mut index = 0usize;
    let mut previous: Option<usize> = None;
    while index < replacements.len() {
        let paragraph_index = replacements[index].paragraph;
        let mut end = index;
        while end < replacements.len() && replacements[end].paragraph == paragraph_index {
            end += 1;
        }
        let group = &replacements[index..end];
        index = end;
        let paragraph = story
            .paragraphs
            .get(paragraph_index)
            .ok_or_else(|| refused(paragraph_index, Refusal::BlockNotFound))?;
        if previous.is_some_and(|last| last >= paragraph_index) {
            return Err(refused(
                paragraph_index,
                Refusal::AmbiguousCompositeSelector,
            ));
        }
        previous = Some(paragraph_index);
        let mut cursor = 0usize;
        let mut text_ranges: Vec<&Replacement> = Vec::new();
        let mut splices = Vec::new();
        let mut new_text = String::with_capacity(paragraph.text.len());
        let mut texts = 0usize;
        let mut fields = 0usize;
        for replacement in group {
            let range = replacement.range.clone();
            if range.is_empty()
                || range.start < cursor
                || range.end > paragraph.text.len()
                || !paragraph.text.is_char_boundary(range.start)
                || !paragraph.text.is_char_boundary(range.end)
            {
                return Err(refused(
                    paragraph_index,
                    Refusal::AmbiguousCompositeSelector,
                ));
            }
            let replaced = replacement.replaced();
            if let Some(text) = replaced {
                validate_text(paragraph_index, text)?;
                total = total
                    .checked_add(text.len())
                    .filter(|total| *total <= MAX_REPLACEMENT_TEXT_BYTES)
                    .ok_or(Error::Limit {
                        kind: crate::document::LimitKind::ReplacementTextBytes,
                        maximum: MAX_REPLACEMENT_TEXT_BYTES,
                        observed: total.saturating_add(text.len()),
                    })?;
            }
            new_text.push_str(&paragraph.text[cursor..range.start]);
            new_text.push_str(replaced.unwrap_or(&paragraph.text[range.clone()]));
            cursor = range.end;
            let field = (!replacement.formatted)
                .then(|| paragraph.fields.iter().find(|field| field.range == range))
                .flatten();
            if let (Some(field), Some(text)) = (field, replaced) {
                if !field.replaceable {
                    return Err(refused(paragraph_index, Refusal::FieldInterior));
                }
                splices.push((
                    field.markup.start,
                    field.markup.end,
                    field_run(&story.xml, field, text),
                ));
                fields += 1;
            } else {
                if !paragraph.is_text(range) {
                    return Err(refused(paragraph_index, Refusal::FieldInterior));
                }
                text_ranges.push(replacement);
                texts += 1;
            }
        }
        new_text.push_str(&paragraph.text[cursor..]);
        let formatted = rewrite_nodes(story, paragraph, &text_ranges, format, &mut splices)?;
        if splices.is_empty() {
            continue;
        }
        splices.sort_by_key(|splice| splice.0);
        if splices.windows(2).any(|pair| pair[0].1 > pair[1].0) {
            return Err(refused(
                paragraph_index,
                Refusal::AmbiguousCompositeSelector,
            ));
        }
        outcome.paragraphs += 1;
        outcome.texts += texts;
        outcome.fields += fields;
        outcome.formatted += formatted;
        changes.push(Change {
            paragraph: paragraph_index,
            splices,
            text: new_text,
        });
    }
    Ok(Plan { changes, outcome })
}

/// The new content of one text node and the ranges of it to isolate.
struct NodeChange {
    piece: usize,
    content: String,
    isolate: Vec<Range<usize>>,
}

/// Rewrite every `w:t` element whose text the replacements change, and
/// split out the runs that formatted replacements isolate. Returns the
/// number of runs formatted.
fn rewrite_nodes<F: FnMut(&mut properties::Run)>(
    story: &Story,
    paragraph: &Paragraph,
    replacements: &[&Replacement],
    format: &mut F,
    splices: &mut Vec<(usize, usize, Vec<u8>)>,
) -> Result<usize> {
    if replacements.is_empty() {
        return Ok(0);
    }
    let xml = story.xml.as_slice();
    let mut changes: Vec<NodeChange> = Vec::new();
    let first = paragraph
        .pieces
        .partition_point(|piece| piece.text.end <= replacements[0].range.start);
    for (offset, piece) in paragraph.pieces[first..].iter().enumerate() {
        if piece.node.is_none() {
            continue;
        }
        let (start, end) = (piece.text.start, piece.text.end);
        let mut overlapping = replacements
            .iter()
            .filter(|replacement| replacement.range.start < end && replacement.range.end > start)
            .peekable();
        if overlapping.peek().is_none() {
            if replacements
                .last()
                .is_some_and(|last| last.range.end <= start)
            {
                break;
            }
            continue;
        }
        let mut content = String::new();
        let mut isolate = Vec::new();
        let mut cursor = start;
        for replacement in overlapping {
            let from = replacement.range.start.max(start);
            if from >= cursor {
                content.push_str(&paragraph.text[cursor..from]);
            }
            let to = replacement.range.end.min(end);
            match replacement.replaced() {
                Some(text) => {
                    if replacement.range.start >= start {
                        if replacement.formatted && !text.is_empty() {
                            isolate.push(content.len()..content.len() + text.len());
                        }
                        content.push_str(text);
                    }
                },
                None => {
                    if replacement.formatted {
                        isolate.push(content.len()..content.len() + (to - from));
                    }
                    content.push_str(&paragraph.text[from..to]);
                },
            }
            cursor = to;
        }
        content.push_str(&paragraph.text[cursor..end]);
        if isolate.is_empty() && content == paragraph.text[start..end] {
            continue;
        }
        changes.push(NodeChange {
            piece: first + offset,
            content,
            isolate,
        });
    }
    let mut formatted = 0usize;
    let mut at = 0usize;
    while at < changes.len() {
        let Some(node) = paragraph.pieces[changes[at].piece].node.as_ref() else {
            at += 1;
            continue;
        };
        let run = node.run;
        let mut next = at;
        while next < changes.len()
            && paragraph.pieces[changes[next].piece]
                .node
                .as_ref()
                .is_some_and(|node| node.run == run)
        {
            next += 1;
        }
        let in_run = &changes[at..next];
        if in_run.iter().all(|change| change.isolate.is_empty()) {
            for change in in_run {
                let Some(node) = &paragraph.pieces[change.piece].node else {
                    continue;
                };
                splices.push((
                    node.element.start,
                    node.element.end,
                    text_elements(xml, node, &change.content),
                ));
            }
        } else {
            let slot = paragraph
                .runs
                .get(run)
                .ok_or_else(|| super::scan::xml_error("a text node lost its run"))?;
            let (bytes, count) = split_run(story, paragraph, slot, run, in_run, format)?;
            splices.push((slot.element.start, slot.element.end, bytes));
            formatted += count;
        }
        at = next;
    }
    Ok(formatted)
}

/// Rewrite one run whose text nodes isolate ranges: each isolated range
/// becomes a run of its own carrying a copy of the run's start tag and
/// properties, formatted by `format`. Returns the bytes replacing the run
/// and the number of runs formatted.
fn split_run<F: FnMut(&mut properties::Run)>(
    story: &Story,
    paragraph: &Paragraph,
    slot: &RunSlot,
    run: usize,
    changes: &[NodeChange],
    format: &mut F,
) -> Result<(Vec<u8>, usize)> {
    let xml = story.xml.as_slice();
    let open = &xml[slot.element.start..slot.open_end];
    let close = &xml[slot.close_start..slot.element.end];
    let properties = slot.properties.clone().map_or(&[][..], |range| &xml[range]);
    let mut cursor = slot
        .properties
        .as_ref()
        .map_or(slot.open_end, |range| range.end);
    let bindings = story
        .bindings
        .get(slot.bindings)
        .ok_or_else(|| super::scan::xml_error("a run lost its namespace scope"))?;
    let mut output = Vec::new();
    let mut current: Vec<u8> = Vec::new();
    let mut formatted = 0usize;
    let mut flush = |current: &mut Vec<u8>, isolated: bool| -> Result<()> {
        if current.is_empty() {
            return Ok(());
        }
        let mut part =
            Vec::with_capacity(open.len() + properties.len() + current.len() + close.len());
        part.extend_from_slice(open);
        part.extend_from_slice(properties);
        part.append(current);
        part.extend_from_slice(close);
        if isolated {
            let before = properties::parse_run(&part, bindings)?;
            let mut target = before.clone();
            format(&mut target);
            if let Some(after) =
                properties::splice_run(&part, &before, &target, bindings).map_err(spliced)?
            {
                part = after;
            }
            formatted += 1;
        }
        output.extend_from_slice(&part);
        Ok(())
    };
    let nodes = paragraph
        .pieces
        .iter()
        .enumerate()
        .filter_map(|(index, piece)| {
            piece
                .node
                .as_ref()
                .filter(|node| node.run == run)
                .map(|node| (index, node))
        });
    for (index, node) in nodes {
        current.extend_from_slice(&xml[cursor..node.element.start]);
        cursor = node.element.end;
        let Some(change) = changes.iter().find(|change| change.piece == index) else {
            current.extend_from_slice(&xml[node.element.clone()]);
            continue;
        };
        let mut at = 0;
        for range in &change.isolate {
            if range.start > at {
                current.extend_from_slice(&text_elements(
                    xml,
                    node,
                    &change.content[at..range.start],
                ));
            }
            flush(&mut current, false)?;
            current.extend_from_slice(&text_elements(xml, node, &change.content[range.clone()]));
            flush(&mut current, true)?;
            at = range.end;
        }
        if at < change.content.len() || change.isolate.is_empty() {
            current.extend_from_slice(&text_elements(xml, node, &change.content[at..]));
        }
    }
    current.extend_from_slice(&xml[cursor..slot.close_start]);
    flush(&mut current, false)?;
    Ok((output, formatted))
}

/// `content` as `w:t` elements with the original start tag, separated by
/// `w:tab` and `w:br` for tabs and line feeds.
pub(super) fn text_elements(xml: &[u8], node: &TextNode, content: &str) -> Vec<u8> {
    let name = &xml[node.name.clone()];
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&b""[..], |colon| &name[..colon]);
    let start_tag = &xml[node.element.start..node.start_tag_end];
    let mut output = Vec::with_capacity(content.len() + 2 * start_tag.len() + 8);
    let mut segments = content.split(['\t', '\n']).peekable();
    let mut separators = content
        .chars()
        .filter(|character| matches!(character, '\t' | '\n'));
    let mut wrote = false;
    while let Some(segment) = segments.next() {
        if !segment.is_empty() || (!wrote && segments.peek().is_none()) {
            write_start_tag(&mut output, start_tag, node.space, segment);
            escape(&mut output, segment);
            output.extend_from_slice(b"</");
            output.extend_from_slice(name);
            output.push(b'>');
            wrote = true;
        }
        if let Some(separator) = separators.next() {
            output.push(b'<');
            output.extend_from_slice(&qualified(
                prefix,
                if separator == '\t' { b"tab" } else { b"br" },
            ));
            output.extend_from_slice(b"/>");
        }
    }
    output
}

fn write_start_tag(output: &mut Vec<u8>, start_tag: &[u8], space: bool, segment: &str) {
    let needs_space = !space
        && (segment.starts_with(char::is_whitespace) || segment.ends_with(char::is_whitespace));
    if needs_space {
        output.extend_from_slice(&start_tag[..start_tag.len() - 1]);
        output.extend_from_slice(b" xml:space=\"preserve\">");
    } else {
        output.extend_from_slice(start_tag);
    }
}

fn escape(output: &mut Vec<u8>, text: &str) {
    for byte in text.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            _ => output.push(byte),
        }
    }
}

/// One run holding `text`, formatted as the field's result.
fn field_run(xml: &[u8], field: &Field, text: &str) -> Vec<u8> {
    if text.is_empty() {
        return Vec::new();
    }
    let run_name = field.run_name.as_slice();
    let prefix = run_name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&b""[..], |colon| &run_name[..colon]);
    let text_name = qualified(prefix, b"t");
    let mut start_tag = Vec::with_capacity(text_name.len() + 2);
    start_tag.push(b'<');
    start_tag.extend_from_slice(&text_name);
    start_tag.push(b'>');
    let mut output = Vec::new();
    output.push(b'<');
    output.extend_from_slice(run_name);
    output.push(b'>');
    if let Some(properties) = &field.properties {
        output.extend_from_slice(&xml[properties.clone()]);
    }
    // The synthetic `w:t` start tag spans `start_tag`; reuse the element
    // writer with a node over a scratch buffer.
    let mut scratch = start_tag.clone();
    scratch.extend_from_slice(b"</");
    scratch.extend_from_slice(&text_name);
    scratch.push(b'>');
    let node = TextNode {
        element: 0..scratch.len(),
        start_tag_end: start_tag.len(),
        name: 1..1 + text_name.len(),
        space: false,
        run: 0,
        group: 0,
        scoped: true,
    };
    output.extend_from_slice(&text_elements(&scratch, &node, text));
    output.extend_from_slice(b"</");
    output.extend_from_slice(run_name);
    output.push(b'>');
    output
}

/// How the offsets of the old story move: the end of every splice and the
/// total length change up to and including it, ascending.
pub(super) struct Shifts {
    ends: Vec<usize>,
    deltas: Vec<isize>,
}

impl Shifts {
    pub(super) fn of(changes: &[Change]) -> Self {
        let mut splices: Vec<(usize, isize)> = changes
            .iter()
            .flat_map(|change| {
                change.splices.iter().map(|(start, end, bytes)| {
                    (
                        *end,
                        isize::try_from(bytes.len()).unwrap_or(isize::MAX)
                            - isize::try_from(end - start).unwrap_or(isize::MAX),
                    )
                })
            })
            .collect();
        splices.sort_unstable_by_key(|splice| splice.0);
        let mut total = 0isize;
        let (ends, deltas) = splices
            .into_iter()
            .map(|(end, delta)| {
                total += delta;
                (end, total)
            })
            .unzip();
        Self { ends, deltas }
    }

    /// Where `offset` (never strictly inside a splice) lies after them.
    pub(super) fn shift(&self, offset: usize) -> usize {
        let before = self.ends.partition_point(|end| *end <= offset);
        let delta = before.checked_sub(1).map_or(0, |at| self.deltas[at]);
        offset.saturating_add_signed(delta)
    }
}
