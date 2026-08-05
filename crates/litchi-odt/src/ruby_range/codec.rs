use super::*;

fn record_range_boundary(
    boundaries: &mut Vec<Option<RangeBoundary>>,
    stack: &[RangeFrame],
    offset: usize,
    prefer_later: bool,
    ruby_epoch: usize,
    inside_ruby: bool,
) {
    if !matches!(
        stack.last(),
        Some(RangeFrame {
            namespace: Ns::Text,
            local,
            ..
        }) if matches!(
            local.as_slice(),
            b"p" | b"h" | b"span" | b"a" | b"meta" | b"meta-field" | b"ruby-base"
        )
    ) || inside_ruby
    {
        return;
    }
    let boundary = RangeBoundary {
        offset,
        container_id: stack.last().unwrap().id,
        ruby_epoch,
    };
    if boundaries.len() <= stack.len() {
        boundaries.resize(stack.len() + 1, None);
    }
    let slot = &mut boundaries[stack.len()];
    if prefer_later || slot.is_none() {
        *slot = Some(boundary);
    }
}

pub(super) fn locate_balanced_ruby_ranges(
    xml: &str,
    paragraph_index: usize,
    range: &Range<usize>,
) -> Result<Vec<Span>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<RangeFrame>::new();
    let mut starts = Vec::<Option<RangeBoundary>>::new();
    let mut ends = Vec::<Option<RangeBoundary>>::new();
    let mut paragraph_count = 0usize;
    let mut target_depth = None;
    let mut text_offset = 0usize;
    let mut previous_end = 0usize;
    let mut next_frame_id = 0usize;
    let mut ruby_epoch = 0usize;
    let mut open_rubies = 0usize;
    let mut events = 0usize;

    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("too many ruby structural range XML events"));
        }
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid ruby structural range XML: {error}")))?;
        let namespace = ns(&resolved);
        let event_end = reader.buffer_position() as usize;
        if target_depth.is_some() {
            if text_offset == range.start {
                record_range_boundary(
                    &mut starts,
                    &stack,
                    previous_end,
                    true,
                    ruby_epoch,
                    open_rubies != 0,
                );
            }
            if text_offset == range.end {
                record_range_boundary(
                    &mut ends,
                    &stack,
                    previous_end,
                    false,
                    ruby_epoch,
                    open_rubies != 0,
                );
            }
        }

        match event {
            Event::Start(ref start) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("ruby structural range XML is too deep"));
                }
                let local = start.local_name().as_ref().to_vec();
                next_frame_id = next_frame_id
                    .checked_add(1)
                    .ok_or_else(|| bad("ruby structural range frame overflow"))?;
                stack.push(RangeFrame {
                    id: next_frame_id,
                    namespace,
                    local: local.clone(),
                });
                if namespace == Ns::Text && local == b"ruby" {
                    ruby_epoch = ruby_epoch
                        .checked_add(1)
                        .ok_or_else(|| bad("ruby structural range epoch overflow"))?;
                    open_rubies = open_rubies
                        .checked_add(1)
                        .ok_or_else(|| bad("ruby structural range nesting overflow"))?;
                }
                if namespace == Ns::Text && local == b"p" {
                    if paragraph_count == paragraph_index {
                        target_depth = Some(stack.len());
                    }
                    paragraph_count = paragraph_count
                        .checked_add(1)
                        .ok_or_else(|| bad("ruby paragraph count overflow"))?;
                }
            },
            Event::Empty(ref start)
                if namespace == Ns::Text && start.local_name().as_ref() == b"p" =>
            {
                paragraph_count = paragraph_count
                    .checked_add(1)
                    .ok_or_else(|| bad("ruby paragraph count overflow"))?;
            },
            Event::Text(ref text) if target_depth.is_some() && open_rubies == 0 => {
                let content = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| bad(format!("invalid ruby structural range text: {error}")))?;
                text_offset = text_offset
                    .checked_add(content.len())
                    .ok_or_else(|| bad("ruby structural range offset overflow"))?;
            },
            Event::CData(ref text) if target_depth.is_some() && open_rubies == 0 => {
                let content = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    bad(format!("invalid ruby structural range CDATA: {error}"))
                })?;
                text_offset = text_offset
                    .checked_add(content.len())
                    .ok_or_else(|| bad("ruby structural range offset overflow"))?;
            },
            Event::GeneralRef(ref reference) if target_depth.is_some() && open_rubies == 0 => {
                let content =
                    crate::elements::xml::decode_reference(reference, "ruby structural range")?;
                text_offset = text_offset
                    .checked_add(content.len())
                    .ok_or_else(|| bad("ruby structural range offset overflow"))?;
            },
            Event::End(_) => {
                let depth = stack.len();
                let frame = stack
                    .pop()
                    .ok_or_else(|| bad("ruby structural range XML depth underflow"))?;
                if frame.namespace == Ns::Text && frame.local == b"ruby" {
                    open_rubies = open_rubies
                        .checked_sub(1)
                        .ok_or_else(|| bad("ruby structural range nesting underflow"))?;
                }
                if target_depth == Some(depth) {
                    target_depth = None;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(bad("DTD and processing instructions are prohibited"));
            },
            Event::Eof => break,
            _ => {},
        }

        if target_depth.is_some() {
            if text_offset == range.start {
                record_range_boundary(
                    &mut starts,
                    &stack,
                    event_end,
                    true,
                    ruby_epoch,
                    open_rubies != 0,
                );
            }
            if text_offset == range.end {
                record_range_boundary(
                    &mut ends,
                    &stack,
                    event_end,
                    false,
                    ruby_epoch,
                    open_rubies != 0,
                );
            }
        }
        previous_end = event_end;
        buffer.clear();
    }

    if !stack.is_empty() {
        return Err(bad("truncated ruby structural range XML"));
    }
    if paragraph_index >= paragraph_count {
        return Err(bad("paragraph index does not exist"));
    }
    if range.end > text_offset {
        return Err(bad("ruby text range is out of bounds"));
    }

    let mut spans = starts
        .into_iter()
        .enumerate()
        .filter_map(|(depth, start)| {
            let start = start?;
            let end = ends.get(depth).copied().flatten()?;
            (start.container_id == end.container_id
                && start.ruby_epoch == end.ruby_epoch
                && start.offset < end.offset)
                .then_some(Span {
                    start: start.offset,
                    end: end.offset,
                })
        })
        .collect::<Vec<_>>();
    spans.sort_by_key(|span| span.end - span.start);
    Ok(spans)
}

#[allow(clippy::too_many_arguments)]
pub(super) fn collect_ruby_text_node(
    xml: &str,
    span: Range<usize>,
    text: &str,
    text_offset: &mut usize,
    range: &Range<usize>,
    stack: &[(Ns, Vec<u8>)],
    fragment: &str,
    target_depth: Option<usize>,
    value: &Annotation,
    pending: &mut Option<PendingRubyRange>,
) -> Result<Option<String>> {
    let Some(target_depth) = target_depth else {
        return Ok(None);
    };
    if stack.len() < target_depth
        || stack
            .iter()
            .any(|(namespace, local)| *namespace == Ns::Text && local == b"ruby")
    {
        return Ok(None);
    }
    let start = *text_offset;
    let end = start
        .checked_add(text.len())
        .ok_or_else(|| bad("ruby text range offset overflow"))?;
    *text_offset = end;
    if range.start >= end || range.end <= start {
        return Ok(None);
    }
    if !ruby_parent(stack.last()) {
        return Err(bad("ruby text range has unsupported inline parent"));
    }
    let local_start = range.start.saturating_sub(start);
    let local_end = range.end.min(end) - start;
    if !text.is_char_boundary(local_start) || !text.is_char_boundary(local_end) {
        return Err(bad("ruby text range is not on a UTF-8 character boundary"));
    }

    let state = if let Some(state) = pending.as_mut() {
        if state.xml_end != span.start || state.stack != stack {
            return Err(bad("ruby text range crosses an inline markup boundary"));
        }
        state
    } else {
        pending.insert(PendingRubyRange {
            xml_start: span.start,
            xml_end: span.start,
            prefix: escape_xml(&text[..local_start]),
            selected: String::new(),
            stack: stack.to_vec(),
        })
    };
    let selected_len = local_end - local_start;
    let total_len = state
        .selected
        .len()
        .checked_add(selected_len)
        .ok_or_else(|| bad("ruby text range size overflow"))?;
    if total_len > MAX_BASE {
        return Err(bad("ruby text range exceeds the base-size limit"));
    }
    state
        .selected
        .try_reserve(selected_len)
        .map_err(|_| bad("ruby text range allocation failed"))?;
    state.selected.push_str(&text[local_start..local_end]);
    state.xml_end = span.end;

    if range.end > end {
        return Ok(None);
    }
    if value.base.xml != escape_xml(&state.selected) {
        return Err(bad(
            "ruby annotation base must equal the selected plain-text range",
        ));
    }
    let suffix = escape_xml(&text[local_end..]);
    let mut output = String::with_capacity(xml.len() + fragment.len());
    output.push_str(&xml[..state.xml_start]);
    output.push_str(&state.prefix);
    output.push_str(fragment);
    output.push_str(&suffix);
    output.push_str(&xml[state.xml_end..]);
    Ok(Some(output))
}
