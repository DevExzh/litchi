//! Paragraph insertion, removal, moves and swaps on producer RTF by source
//! splicing (records 0293, 0308).
//!
//! The lifecycle projection and the ordinary paragraph map accept only plain
//! ASCII bodies with default formatting, so producer RTF (Word, LibreOffice,
//! WordPad: `\pard` paragraph properties, character-formatting groups,
//! fonts and colours) was refused. This seam instead cuts the body into the
//! source blocks its root-level `\par` and `\row` controls delimit
//! (paragraph segments, and table rows that keep their place among the
//! paragraphs that do not move) and inserts, removes, moves or exchanges
//! whole paragraph segments. Every byte outside the changed segments is
//! copied verbatim, so a move and its inverse move restore the source bytes
//! where nothing below applies.
//!
//! Three producer shapes need more than a cut, and their results are
//! value-exact rather than byte-exact:
//!
//! - a formatting group that holds several paragraphs
//!   (`{\f1 one\par two\par}`) is first rewritten as one group per
//!   paragraph, each later group replaying the control words set before it;
//! - a paragraph that relied on root-level state its old neighbour set gets
//!   that state replayed (the control words since the last `\pard` and
//!   `\plain`) wherever its predecessor changed; table rows always do;
//! - an inserted paragraph is a group, `{\pard\plain text\par}`, so the
//!   blocks after it keep the state they had.
//!
//! RTF state is positional and anchored content (fields, pictures,
//! bookmarks) can span segments. Rather than model those rules, the
//! candidate is reparsed with the source's open policy and must read back as
//! the expected paragraph list, each paragraph with its exact text, paragraph
//! properties and runs, every table unchanged, and every other collection
//! unchanged (or, for a removal, smaller by what the removed segment held,
//! with no bookmark or comment range cut). Anything else is a typed refusal
//! that names what moved.
//!
//! A root-level section break (`\sect` and the section controls after it)
//! is a block of its own that keeps its place among the paragraphs, like a
//! table row; every section must read back with its properties and headers
//! unchanged (record 0328).

use std::ops::Range;

use bumpalo::Bump;

use super::{Commit, Error, Operation, Package, semantic_changes};
use crate::lexer::{ControlWord, Lexer, Token};

/// One body block: a paragraph segment or a table row.
struct Block {
    range: Range<usize>,
    paragraph: bool,
}

/// Whether a root-level control word resets paragraph or character state.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Reset {
    None,
    Paragraph,
    Character,
}

const GROUP_CONTINUES: &str = "a paragraph ends inside a group that continues with more content";

/// One paragraph's source segment: from the end of the previous paragraph's
/// boundary through the end of its own `\par` (and the delimiter space that
/// belongs to that control word).
struct Segments {
    ranges: Vec<Range<usize>>,
    /// Every body block in source order: paragraph segments and table rows
    /// (`\trowd` or `\intbl` through `\row`), which are not body
    /// paragraphs and keep their place among the paragraphs that do not
    /// move.
    blocks: Vec<Block>,
    /// The block of each paragraph.
    paragraph_blocks: Vec<usize>,
    /// Where the body holds control words the lexer does not know.
    unknown: Vec<usize>,
    /// The root-level control words that set paragraph or character state,
    /// in source order, with whether each is `\pard` or `\plain`.
    controls: Vec<(Range<usize>, Reset)>,
    /// Whether the last paragraph has no `\par` of its own: its segment ends
    /// before the root group's closing brace.
    unterminated_last: bool,
}

/// One paragraph of the expected result: a source paragraph, or the
/// paragraph an insertion authors.
#[derive(Clone, Copy)]
enum Slot<'a> {
    Source(usize),
    New(&'a str),
}

/// Publishes one paragraph insertion, removal, move or swap on a producer
/// source.
pub(super) fn commit(source: Package, operation: &Operation) -> Result<Commit, Error> {
    super::ensure_changed_publication_allowed(&source)?;
    let count = source.paragraph_count();
    let (order, touched) = match operation {
        Operation::RestoreParagraph { position, text } => {
            if *position > count {
                return Err(Error::ParagraphOutOfRange {
                    position: *position,
                    count,
                });
            }
            let mut order: Vec<Slot<'_>> = (0..count).map(Slot::Source).collect();
            order.insert(*position, Slot::New(text));
            (order, Vec::new())
        },
        Operation::RemoveParagraph { position, text } => {
            check_text(&source, *position, text)?;
            let order = (0..count)
                .filter(|index| index != position)
                .map(Slot::Source)
                .collect();
            (order, vec![*position])
        },
        Operation::MoveParagraph {
            position,
            final_position,
            text,
        } => {
            check_text(&source, *position, text)?;
            let mut order: Vec<usize> = (0..count).filter(|index| index != position).collect();
            if *final_position > order.len() {
                return Err(Error::ParagraphOutOfRange {
                    position: *final_position,
                    count,
                });
            }
            order.insert(*final_position, *position);
            (
                order.into_iter().map(Slot::Source).collect(),
                vec![*position, *final_position],
            )
        },
        Operation::SwapParagraph {
            first,
            second,
            first_text,
            second_text,
        } => {
            check_text(&source, *first, first_text)?;
            check_text(&source, *second, second_text)?;
            let mut order: Vec<usize> = (0..count).collect();
            order.swap(*first, *second);
            (
                order.into_iter().map(Slot::Source).collect(),
                vec![*first, *second],
            )
        },
        _ => {
            return Err(Error::unsupported_operation(
                "body splices and retained-destination edits cannot compose safely",
            ));
        },
    };
    let source_bytes = source.source_bytes().ok_or(Error::unsupported_operation(
        "snapshot has no exact RTF source",
    ))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::unsupported_source(
            "producer paragraph splicing refuses compressed RTF",
        ));
    }
    let input = std::str::from_utf8(source_bytes).map_err(|_error| {
        Error::unsupported_source(
            "producer paragraph splicing requires an ASCII or UTF-8 transport",
        )
    })?;
    // A formatting group that holds several paragraphs is first rewritten
    // as one group per paragraph (the result is then value-exact, not
    // byte-exact); the readback below proves the rewrite kept every run.
    let normalized;
    let (source_bytes, cut) = match segments(&source, input) {
        Ok(cut) => (source_bytes, cut),
        Err(error @ Error::Refused(refusal))
            if refusal.kind() == super::RefusalKind::UnsupportedContent
                && refusal.reason() == GROUP_CONTINUES =>
        {
            normalized = split_paragraph_groups(&source, input)?.ok_or(error)?;
            let cut = segments(&source, &normalized)?;
            (normalized.as_bytes(), cut)
        },
        Err(error) => return Err(error),
    };
    if cut.ranges.len() != count {
        return Err(Error::unsupported_source(
            "the body's paragraphs are not delimited by root-level \\par controls",
        ));
    }
    let last = count.saturating_sub(1);
    if cut.unterminated_last
        && touched.contains(&last)
        && !matches!(operation, Operation::RemoveParagraph { .. })
    {
        return Err(Error::unsupported_source(
            "the last paragraph has no \\par of its own and cannot move",
        ));
    }
    // The exact splice first; when a paragraph relied on state its old
    // neighbours set at the root, the splice again with that state replayed
    // where the order changed (value-exact).
    let removed = match operation {
        Operation::RemoveParagraph { position, .. } => {
            let range = cut.ranges.get(*position).cloned().unwrap_or_default();
            if cut.unknown.iter().any(|offset| range.contains(offset)) {
                return Err(Error::unsupported_source(
                    "the removed paragraph holds control words this reader does not know",
                ));
            }
            source_bytes.get(range)
        },
        _ => None,
    };
    let attempt = |replay: bool| -> Result<(Vec<u8>, Package), Error> {
        let bytes = splice(
            source_bytes,
            &cut,
            operation,
            &order,
            replay,
            source.limits(),
        )?;
        let snapshot =
            Package::from_bytes_with_options(&bytes, source.options()).map_err(|_error| {
                Error::unsupported_source("the spliced paragraph segments do not parse as RTF")
            })?;
        verify(&source, &snapshot, &order, removed)?;
        Ok((bytes, snapshot))
    };
    let (bytes, snapshot) = match attempt(false) {
        Ok(done) => done,
        Err(first @ Error::Refused(refusal))
            if refusal.kind() == super::RefusalKind::UnsupportedContent
                && !cut.controls.is_empty() =>
        {
            attempt(true).map_err(|_second| first)?
        },
        Err(error) => return Err(error),
    };
    // The inverse move or swap must be plannable on the candidate the same
    // way, so the edit is reversible through this seam too. (A removal's
    // inverse is an insertion, which this seam does not plan.)
    if matches!(
        operation,
        Operation::RemoveParagraph { .. } | Operation::RestoreParagraph { .. }
    ) {
        let delta = semantic_changes(std::slice::from_ref(operation), &[]);
        return Ok(Commit::new(source, snapshot, 1, delta));
    }
    let candidate_input = std::str::from_utf8(&bytes).map_err(|_error| {
        Error::unsupported_source("the spliced paragraph segments are not UTF-8")
    })?;
    if segments(&snapshot, candidate_input)?.ranges.len() != snapshot.paragraph_count() {
        return Err(Error::unsupported_source(
            "the spliced paragraphs would not be delimited by root-level \\par controls",
        ));
    }
    let delta = semantic_changes(std::slice::from_ref(operation), &[]);
    Ok(Commit::new(source, snapshot, 1, delta))
}

fn check_text(source: &Package, position: usize, text: &str) -> Result<(), Error> {
    if super::paragraph_text_at(source, position)? == text {
        Ok(())
    } else {
        Err(Error::StalePrecondition("paragraph text differs"))
    }
}

/// The source segment of each body paragraph (record 0328).
pub(super) fn paragraph_ranges(source: &Package, input: &str) -> Result<Vec<Range<usize>>, Error> {
    Ok(segments(source, input)?.ranges)
}

/// Cuts the body into one source segment per paragraph.
fn segments(source: &Package, input: &str) -> Result<Segments, Error> {
    let arena = Bump::new();
    let mut lexer = Lexer::new_with_limits(input, &arena, source.limits())
        .for_document(source.options().leniency_value());
    let (tokens, spans) = lexer.tokenize_with_spans()?;
    let bytes = input.as_bytes();
    let mut depth = 0_usize;
    let mut boundaries = Vec::new();
    let mut first_start = None;
    let mut root_end = None;
    // Per open group: whether it, or a group around it, is a destination
    // (header, note, field, picture, `\*` extension) whose `\par` controls
    // belong to another story.
    let mut destinations: Vec<bool> = Vec::new();
    // A `\par` that ends a character-formatting group (`{\b text\par}`,
    // as Word and WordPad write) bounds its paragraph where the group
    // closes back to the root; content after it inside the group does not.
    let mut pending_group_par = false;
    // Whether the pending group boundary ends a table row.
    let mut pending_row = false;
    let mut first_fallback = None;
    // The generated list marker group (`{\listtext ...}`, `{\pntext ...}`)
    // that precedes a paragraph belongs to it: its start, and its end once
    // closed, while nothing else followed at the root.
    let mut marker: Option<(usize, Option<usize>)> = None;
    let mut controls = Vec::new();
    // Inside a table row: from `\trowd` or `\intbl` at the root through
    // its `\row`, whose cell `\par`/`\cell` controls are not body
    // paragraph boundaries. Each boundary records whether it ends a row.
    let mut in_row = false;
    let mut rows: Vec<bool> = Vec::new();
    let mut unknown = Vec::new();
    // A root-level `\sect` and the section controls after it, through the
    // next paragraph reset, text or group, are a block of their own that is
    // not a body paragraph (record 0328, the follow-up of record 0308).
    let mut in_section = false;
    for (token, span) in tokens.iter().zip(&spans) {
        if in_section {
            match token {
                Token::Control(control) if depth == 1 && !matches!(control, ControlWord::Pard) => {
                    continue;
                },
                _ => {
                    boundaries.push(span.start);
                    rows.push(true);
                    in_section = false;
                },
            }
        }
        if depth == 1 {
            match token {
                Token::OpenBrace => {
                    let after = bytes.get(span.end..).unwrap_or_default().trim_ascii_start();
                    marker = (after.starts_with(br"\listtext") || after.starts_with(br"\pntext"))
                        .then_some((span.start, None));
                },
                Token::Control(ControlWord::Pard) => {},
                Token::Text(text) if text.trim().is_empty() => {},
                _ => marker = None,
            }
        }
        if depth == 1
            && matches!(token, Token::Control(_))
            && !destinations.last().copied().unwrap_or(false)
        {
            let raw = bytes.get(span.clone()).unwrap_or_default();
            if sets_state(raw) {
                let reset = match control_name(raw) {
                    b"pard" => Reset::Paragraph,
                    b"plain" => Reset::Character,
                    _ => Reset::None,
                };
                controls.push((span.clone(), reset));
            }
        }
        if matches!(token, Token::Control(ControlWord::Unknown(..)))
            && !destinations.last().copied().unwrap_or(false)
        {
            unknown.push(span.start);
        }
        if pending_group_par {
            match token {
                Token::CloseBrace => {
                    depth = depth.saturating_sub(1);
                    destinations.pop();
                    if depth == 1 {
                        boundaries.push(boundary_end(input, span.end));
                        rows.push(pending_row);
                        if pending_row {
                            in_row = false;
                        }
                        pending_group_par = false;
                        pending_row = false;
                    }
                    continue;
                },
                Token::Text(text) if text.trim().is_empty() => continue,
                _ => return Err(Error::unsupported_source(GROUP_CONTINUES)),
            }
        }
        let in_destination = destinations.last().copied().unwrap_or(false);
        match token {
            Token::OpenBrace => {
                let destination =
                    in_destination || opens_destination(bytes.get(span.end..).unwrap_or_default());
                if depth == 1 && !destination && first_start.is_none() {
                    // A group that opens with a paragraph reset is a
                    // paragraph of its own (an inserted one is written so).
                    let after = bytes.get(span.end..).unwrap_or_default().trim_ascii_start();
                    if boundaries.is_empty()
                        && after.starts_with(br"\pard")
                        && !after.get(5).is_some_and(u8::is_ascii_alphabetic)
                    {
                        first_start = Some(span.start);
                    } else {
                        first_fallback.get_or_insert(span.start);
                    }
                }
                depth += 1;
                destinations.push(destination);
            },
            Token::CloseBrace => {
                depth = depth.saturating_sub(1);
                destinations.pop();
                if depth == 0 && root_end.is_none() {
                    root_end = Some(span.start);
                }
                if depth == 1
                    && let Some((_, end @ None)) = &mut marker
                {
                    *end = Some(span.end);
                }
            },
            Token::Control(ControlWord::TableRowDefaults | ControlWord::InTable)
                if depth == 1 && !in_destination && !in_row =>
            {
                in_row = true;
                if boundaries.is_empty() && first_start.is_none() {
                    first_start = Some(match marker {
                        Some((start, Some(_))) => start,
                        _ => first_fallback.unwrap_or(span.start),
                    });
                }
            },
            Token::Control(ControlWord::TableRow) if depth == 1 && in_row => {
                boundaries.push(boundary_end(input, span.end));
                rows.push(true);
                in_row = false;
            },
            Token::Control(ControlWord::TableRow) if depth > 1 && in_row && !in_destination => {
                // A row that ends a formatting group (`{\b ...\row }`, as
                // Word writes it) ends where the group closes.
                pending_group_par = true;
                pending_row = true;
            },
            Token::Control(ControlWord::Par | ControlWord::TableCell) if in_row => {},
            Token::Control(ControlWord::Pard) if depth == 1 && boundaries.is_empty() && !in_row => {
                // The last paragraph reset before the first boundary opens
                // the first paragraph; header destinations precede it, and
                // so does a list marker group just before the reset.
                first_start = Some(match marker {
                    Some((start, Some(_))) => start,
                    _ => span.start,
                });
            },
            Token::Control(ControlWord::Par) if depth == 1 => {
                if boundaries.is_empty() && first_start.is_none() && first_fallback.is_none() {
                    // An empty first paragraph is its `\par` alone.
                    first_fallback = Some(span.start);
                }
                boundaries.push(boundary_end(input, span.end));
                rows.push(false);
            },
            Token::Text(text) if depth == 1 && first_start.is_none() && !text.trim().is_empty() => {
                // Without a `\pard`, the first paragraph starts at its first
                // text (`{\rtf1 1\par 2\par}`); controls before it stay
                // with the document. Readback refuses a paragraph that needs
                // them.
                first_fallback.get_or_insert(span.start);
            },
            Token::Control(ControlWord::Par) if depth > 1 && !in_destination && !in_row => {
                pending_group_par = true;
            },
            Token::Control(ControlWord::TableRow | ControlWord::TableCell) if !in_destination => {
                return Err(Error::unsupported_source(
                    "producer paragraph splicing refuses table rows that do not open and close at the body's root",
                ));
            },
            Token::Control(ControlWord::SectionBreak)
                if depth == 1 && !in_destination && !in_row && !pending_group_par =>
            {
                let start =
                    boundaries
                        .last()
                        .copied()
                        .or(first_start)
                        .ok_or(Error::unsupported_source(
                            "a section break precedes the body's first paragraph",
                        ))?;
                // `\sect` ends the paragraph it follows when no `\par` did.
                if tail_has_content(bytes.get(start..span.start).unwrap_or_default()) {
                    boundaries.push(span.start);
                    rows.push(false);
                }
                in_section = true;
            },
            Token::Control(ControlWord::SectionBreak) if !in_destination => {
                return Err(Error::unsupported_source(
                    "producer paragraph splicing refuses section breaks inside groups or table rows",
                ));
            },
            _ => {},
        }
    }
    let first_start = first_start
        .or(first_fallback)
        .ok_or(Error::unsupported_source(
            "the first paragraph has no root-level \\pard or text that proves where it starts",
        ))?;
    let root_end = root_end.ok_or(Error::unsupported_source(
        "the RTF root group is not closed",
    ))?;
    // Document-level controls before the body are in effect everywhere.
    controls.retain(|(range, _)| range.start >= first_start);
    if in_row {
        return Err(Error::unsupported_source(
            "a table row of the body has no \\row of its own",
        ));
    }
    let mut blocks = Vec::new();
    blocks
        .try_reserve_exact(boundaries.len() + 1)
        .map_err(|_error| Error::Write("could not reserve paragraph segments".to_string()))?;
    let mut start = first_start;
    for (end, row) in boundaries.into_iter().zip(rows) {
        if end <= start {
            return Err(Error::unsupported_source(
                "a root-level \\par precedes the first paragraph's start",
            ));
        }
        blocks.push(Block {
            range: start..end,
            paragraph: !row,
        });
        start = end;
    }
    // Text after the last boundary is an unterminated last paragraph.
    let tail = bytes.get(start..root_end).unwrap_or_default();
    let unterminated_last = tail_has_content(tail);
    if unterminated_last {
        blocks.push(Block {
            range: start..root_end,
            paragraph: true,
        });
    }
    let paragraph_blocks: Vec<usize> = blocks
        .iter()
        .enumerate()
        .filter_map(|(index, block)| block.paragraph.then_some(index))
        .collect();
    let ranges = paragraph_blocks
        .iter()
        .filter_map(|&index| blocks.get(index).map(|block| block.range.clone()))
        .collect();
    Ok(Segments {
        ranges,
        blocks,
        paragraph_blocks,
        unknown,
        controls,
        unterminated_last,
    })
}

/// One root-level formatting group that holds a paragraph boundary
/// followed by more content.
struct SplitGroup {
    /// The group, `{` through `}`.
    range: Range<usize>,
    /// The spans of the group's own state-setting control words.
    controls: Vec<Range<usize>>,
    /// Each paragraph boundary end inside the group, with the number of
    /// `controls` before it.
    cuts: Vec<(usize, usize)>,
}

/// Rewrites every root-level formatting group that holds a `\par`
/// followed by more content (`{\f1 one\par two\par}`) as one group per
/// paragraph (`{\f1 one\par}{\f1 two\par}`). Each later group replays
/// the control words the source group set before it, so every paragraph
/// starts in the state it had. `None` when no group needs it.
fn split_paragraph_groups(source: &Package, input: &str) -> Result<Option<String>, Error> {
    let arena = Bump::new();
    let mut lexer = Lexer::new_with_limits(input, &arena, source.limits())
        .for_document(source.options().leniency_value());
    let (tokens, spans) = lexer.tokenize_with_spans()?;
    let bytes = input.as_bytes();
    let mut depth = 0_usize;
    let mut destinations: Vec<bool> = Vec::new();
    let mut open: Option<SplitGroup> = None;
    let mut after_par = false;
    let mut continued = false;
    let mut groups: Vec<SplitGroup> = Vec::new();
    for (token, span) in tokens.iter().zip(&spans) {
        let in_destination = destinations.last().copied().unwrap_or(false);
        match token {
            Token::OpenBrace => {
                let destination =
                    in_destination || opens_destination(bytes.get(span.end..).unwrap_or_default());
                if depth == 1 && !destination {
                    open = Some(SplitGroup {
                        range: span.start..span.start,
                        controls: Vec::new(),
                        cuts: Vec::new(),
                    });
                    after_par = false;
                    continued = false;
                } else if depth == 2 && after_par {
                    continued = true;
                }
                depth += 1;
                destinations.push(destination);
            },
            Token::CloseBrace => {
                depth = depth.saturating_sub(1);
                destinations.pop();
                if depth == 1
                    && let Some(mut group) = open.take()
                {
                    group.range.end = span.end;
                    if group.cuts.len() > 1 || (!group.cuts.is_empty() && continued) {
                        groups.push(group);
                    }
                }
            },
            Token::Control(ControlWord::Par) if depth == 2 && open.is_some() => {
                if let Some(group) = &mut open {
                    group
                        .cuts
                        .push((boundary_end(input, span.end), group.controls.len()));
                }
                after_par = true;
                continued = false;
            },
            Token::Control(_) if depth == 2 && !in_destination => {
                if let Some(group) = &mut open {
                    if after_par {
                        continued = true;
                    }
                    if sets_state(bytes.get(span.clone()).unwrap_or_default()) {
                        group.controls.push(span.clone());
                    }
                }
            },
            Token::Text(text) if depth == 2 && after_par && !text.trim().is_empty() => {
                continued = true;
            },
            Token::Binary(_) if depth == 2 && after_par => continued = true,
            _ => {},
        }
    }
    if groups.is_empty() {
        return Ok(None);
    }
    let mut output = String::new();
    output
        .try_reserve(input.len().saturating_add(groups.len() * 64))
        .map_err(|_error| Error::Write("could not reserve the split RTF groups".to_string()))?;
    let mut cursor = 0;
    for group in &groups {
        let text = |range: Range<usize>| {
            input.get(range).ok_or(Error::unsupported_source(
                "a paragraph group is outside the source",
            ))
        };
        output.push_str(text(cursor..group.range.start)?);
        let body_end = group.range.end.saturating_sub(1);
        let mut piece_start = group.range.start + 1;
        let mut replay = 0;
        let mut first = true;
        for &(cut, controls) in &group.cuts {
            let piece = text(piece_start..cut)?;
            let replayed = group.controls.get(..replay).unwrap_or_default();
            push_piece(&mut output, input, piece, replayed, first);
            first = false;
            piece_start = cut;
            replay = controls;
        }
        let rest = text(piece_start..body_end)?;
        if rest.trim().is_empty() {
            output.push_str(rest);
        } else {
            let replayed = group.controls.get(..replay).unwrap_or_default();
            push_piece(&mut output, input, rest, replayed, false);
        }
        cursor = group.range.end;
    }
    output.push_str(input.get(cursor..).unwrap_or_default());
    if output.len() > source.limits().max_input_bytes() {
        return Err(Error::InputTooLarge {
            observed: output.len(),
            limit: source.limits().max_input_bytes(),
        });
    }
    Ok(Some(output))
}

/// Appends one paragraph's piece of a split group as its own group, after
/// the replayed controls. A list marker group (`{\listtext ...}`) that opens
/// a later piece stays at the root, before the group, as producers write
/// it for an ordinary paragraph.
fn push_piece(
    output: &mut String,
    input: &str,
    piece: &str,
    controls: &[Range<usize>],
    first: bool,
) {
    let mut piece = piece;
    if !first {
        let trimmed = piece.trim_start();
        let leading = piece.get(..piece.len() - trimmed.len()).unwrap_or_default();
        if (trimmed.starts_with("{\\listtext") || trimmed.starts_with("{\\pntext"))
            && let Some(end) = group_end(trimmed.as_bytes())
            && let (Some(marker), Some(after)) = (trimmed.get(..end), trimmed.get(end..))
        {
            output.push_str(leading);
            output.push_str(marker);
            piece = after;
            if piece.trim().is_empty() {
                output.push_str(piece);
                return;
            }
        }
    }
    output.push('{');
    if !first {
        // Replayed controls are within the source's limits by construction.
        let _ = push_replay(output, input, controls);
    }
    output.push_str(piece);
    output.push('}');
}

/// The end of the group that opens at the start of `bytes`.
fn group_end(bytes: &[u8]) -> Option<usize> {
    let mut depth = 0_usize;
    let mut index = 0;
    while let Some(&byte) = bytes.get(index) {
        match byte {
            b'\\' => index += 1,
            b'{' => depth += 1,
            b'}' => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    return Some(index + 1);
                }
            },
            _ => {},
        }
        index += 1;
    }
    None
}

/// Appends the control words at `controls`, each delimited by one space.
fn push_replay(output: &mut String, input: &str, controls: &[Range<usize>]) -> Result<(), Error> {
    for range in controls {
        let raw = input.get(range.clone()).ok_or(Error::unsupported_source(
            "a replayed control word is outside the source",
        ))?;
        output.push_str(raw.strip_suffix(' ').unwrap_or(raw));
        output.push(' ');
    }
    Ok(())
}

/// The letters of the control word at `raw` (empty for a control symbol).
fn control_name(raw: &[u8]) -> &[u8] {
    let Some(rest) = raw.strip_prefix(b"\\") else {
        return &[];
    };
    let letters = rest
        .iter()
        .take_while(|byte| byte.is_ascii_alphabetic())
        .count();
    rest.get(..letters).unwrap_or_default()
}

/// Whether the control word at `raw` sets group state (character or
/// paragraph formatting) rather than producing text or a break.
fn sets_state(raw: &[u8]) -> bool {
    let word = control_name(raw);
    !word.is_empty()
        && !matches!(
            word,
            b"par"
                | b"line"
                | b"lbr"
                | b"tab"
                | b"u"
                | b"bin"
                | b"page"
                | b"column"
                | b"sect"
                | b"cell"
                | b"row"
                | b"nestcell"
                | b"nestrow"
                | b"emdash"
                | b"endash"
                | b"emspace"
                | b"enspace"
                | b"qmspace"
                | b"bullet"
                | b"lquote"
                | b"rquote"
                | b"ldblquote"
                | b"rdblquote"
                | b"zwj"
                | b"zwnj"
                | b"zwbo"
                | b"zwnbo"
                | b"ltrmark"
                | b"rtlmark"
                | b"chdate"
                | b"chdpl"
                | b"chdpa"
                | b"chtime"
                | b"chpgn"
                | b"sectnum"
                | b"chftn"
                | b"chatn"
                | b"chftnsep"
                | b"chftnsepc"
        )
}

/// The end of a paragraph boundary whose `\par` (or closing brace) ends at
/// `end`: the space delimiting the control word and the line breaks after it
/// belong to the boundary, so a segment moved elsewhere carries them and a
/// move and its inverse restore the source bytes.
fn boundary_end(input: &str, end: usize) -> usize {
    let bytes = input.as_bytes();
    let mut end = end;
    if bytes.get(end) == Some(&b' ')
        && end.checked_sub(1).and_then(|before| bytes.get(before)) != Some(&b' ')
    {
        end += 1;
    }
    while matches!(bytes.get(end), Some(b'\r' | b'\n')) {
        end += 1;
    }
    end
}

/// Whether the group whose `{` precedes `after` is a destination: an
/// ignorable `\*` destination or one whose `\par` controls belong to
/// another story or to no story.
fn opens_destination(after: &[u8]) -> bool {
    let after = after.trim_ascii_start();
    let Some(rest) = after.strip_prefix(b"\\") else {
        return false;
    };
    if rest.starts_with(b"*") {
        return true;
    }
    let word: Vec<u8> = rest
        .iter()
        .take_while(|byte| byte.is_ascii_alphabetic())
        .copied()
        .collect();
    matches!(
        word.as_slice(),
        b"header"
            | b"headerl"
            | b"headerr"
            | b"headerf"
            | b"footer"
            | b"footerl"
            | b"footerr"
            | b"footerf"
            | b"footnote"
            | b"annotation"
            | b"atnid"
            | b"pict"
            | b"object"
            | b"shp"
            | b"shpgrp"
            | b"shptxt"
            | b"field"
            | b"fldinst"
            | b"fldrslt"
            | b"info"
            | b"fonttbl"
            | b"colortbl"
            | b"stylesheet"
            | b"listtable"
            | b"listoverridetable"
            | b"revtbl"
            | b"nonshppict"
            | b"shppict"
            | b"txe"
            | b"xe"
            | b"tc"
            | b"bkmkstart"
            | b"bkmkend"
            | b"do"
            | b"result"
            | b"mmathPr"
            | b"rsidtbl"
            | b"pgdsctbl"
            | b"filetbl"
            | b"latentstyles"
            | b"themedata"
            | b"colorschememapping"
            | b"datastore"
            | b"xmlnstbl"
            | b"wgrffmtfilter"
            | b"defchp"
            | b"defpap"
    )
}

/// Whether the source between the last `\par` and the root's closing brace
/// holds paragraph content rather than whitespace or trailing groups.
fn tail_has_content(tail: &[u8]) -> bool {
    let mut depth = 0_usize;
    let mut index = 0;
    while let Some(&byte) = tail.get(index) {
        match byte {
            b'{' => depth += 1,
            b'}' => depth = depth.saturating_sub(1),
            b'\\' if depth == 0 => {
                let word: Vec<u8> = tail
                    .get(index + 1..)
                    .unwrap_or_default()
                    .iter()
                    .take_while(|byte| byte.is_ascii_alphabetic())
                    .copied()
                    .collect();
                // Paragraph-property resets and section ends carry no text.
                if !matches!(word.as_slice(), b"pard" | b"plain" | b"sectd" | b"par") {
                    return true;
                }
                index += word.len();
            },
            b'\r' | b'\n' | b' ' => {},
            _ if depth == 0 => return true,
            _ => {},
        }
        index += 1;
    }
    false
}

fn part(source: &[u8], range: Range<usize>) -> Result<&[u8], Error> {
    source.get(range).ok_or(Error::unsupported_source(
        "paragraph segment is outside the source",
    ))
}

/// The root-level state in effect at `offset`: the state-setting control
/// words since the last point where both a `\pard` and a `\plain` were
/// seen (or since the body start), each delimited by one space.
fn replay(source: &[u8], cut: &Segments, offset: usize, output: &mut Vec<u8>) -> Result<(), Error> {
    let before = cut
        .controls
        .iter()
        .take_while(|(range, _)| range.end <= offset)
        .count();
    let controls = cut.controls.get(..before).unwrap_or_default();
    let last = |kind: Reset| controls.iter().rposition(|(_, reset)| *reset == kind);
    let start = match (last(Reset::Paragraph), last(Reset::Character)) {
        (Some(paragraph), Some(character)) => paragraph.min(character),
        _ => 0,
    };
    for (range, _) in controls.get(start..).unwrap_or_default() {
        let raw = part(source, range.clone())?;
        output.extend_from_slice(raw.strip_suffix(b" ").unwrap_or(raw));
        output.push(b' ');
    }
    Ok(())
}

/// One block of the assembled body.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Item<'a> {
    Block(usize),
    New(&'a str),
}

/// The body's blocks after `operation`: table rows keep their place among
/// the paragraphs that do not move.
fn arrange<'a>(
    cut: &Segments,
    operation: &'a Operation,
    order: &[Slot<'_>],
) -> Result<Vec<Item<'a>>, Error> {
    let missing = || Error::unsupported_source("paragraph segment is missing");
    let block_of = |paragraph: usize| {
        cut.paragraph_blocks
            .get(paragraph)
            .copied()
            .ok_or_else(missing)
    };
    let mut items: Vec<Item<'a>> = (0..cut.blocks.len()).map(Item::Block).collect();
    let position_of = |items: &[Item<'a>], block: usize| {
        items
            .iter()
            .position(|item| *item == Item::Block(block))
            .ok_or_else(missing)
    };
    match operation {
        Operation::RestoreParagraph { position, text } => {
            let at = match cut.paragraph_blocks.get(*position) {
                Some(&block) => position_of(&items, block)?,
                None => items.len(),
            };
            items.insert(at, Item::New(text));
        },
        Operation::RemoveParagraph { position, .. } => {
            let at = position_of(&items, block_of(*position)?)?;
            items.remove(at);
        },
        Operation::MoveParagraph {
            position,
            final_position,
            ..
        } => {
            let moved = block_of(*position)?;
            let at = position_of(&items, moved)?;
            items.remove(at);
            let at = if *final_position == 0 {
                let Some(Slot::Source(next)) = order.get(1) else {
                    return Err(missing());
                };
                position_of(&items, block_of(*next)?)?
            } else {
                let Some(Slot::Source(previous)) = order.get(final_position - 1) else {
                    return Err(missing());
                };
                position_of(&items, block_of(*previous)?)? + 1
            };
            items.insert(at, Item::Block(moved));
        },
        Operation::SwapParagraph { first, second, .. } => {
            let left = position_of(&items, block_of(*first)?)?;
            let right = position_of(&items, block_of(*second)?)?;
            items.swap(left, right);
        },
        _ => {
            return Err(Error::unsupported_operation(
                "body splices and retained-destination edits cannot compose safely",
            ));
        },
    }
    Ok(items)
}

/// Assembles the body after `operation`: the bytes before the first block,
/// each block (or an inserted paragraph), and the bytes after the last. A
/// table row whose predecessor changed is preceded by the root-level state
/// its source predecessor left; with `replay`, so is a paragraph.
fn splice(
    source: &[u8],
    cut: &Segments,
    operation: &Operation,
    order: &[Slot<'_>],
    replay_state: bool,
    limits: crate::SourceLimits,
) -> Result<Vec<u8>, Error> {
    let blocks = cut.blocks.as_slice();
    let (Some(first), Some(last)) = (blocks.first(), blocks.last()) else {
        return Err(Error::unsupported_source(
            "producer paragraph splicing needs a paragraph to anchor to",
        ));
    };
    let items = arrange(cut, operation, order)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len().saturating_add(64))
        .map_err(|_error| Error::Write("could not reserve the spliced RTF".to_string()))?;
    output.extend_from_slice(part(source, 0..first.range.start)?);
    let mut previous: Option<usize> = None;
    for (index, item) in items.iter().enumerate() {
        match item {
            Item::New(text) => {
                if index + 1 == items.len() && cut.unterminated_last {
                    return Err(Error::unsupported_source(
                        "a paragraph cannot be appended after a last paragraph that has no \\par of its own",
                    ));
                }
                // A group scopes the new paragraph's reset properties, so
                // the blocks after it keep the state they had.
                output.extend_from_slice(br"{\pard\plain ");
                super::destination_splice::encode_text(text, false, &mut output)?;
                output.extend_from_slice(br"\par}");
            },
            Item::Block(block) => {
                let current = blocks
                    .get(*block)
                    .ok_or(Error::unsupported_source("paragraph segment is missing"))?;
                let range = &current.range;
                if (replay_state || !current.paragraph) && previous != block.checked_sub(1) {
                    replay(source, cut, range.start, &mut output)?;
                }
                output.extend_from_slice(part(source, range.clone())?);
                previous = Some(*block);
            },
        }
    }
    output.extend_from_slice(part(source, last.range.end..source.len())?);
    if output.len() > limits.max_input_bytes() {
        return Err(Error::InputTooLarge {
            observed: output.len(),
            limit: limits.max_input_bytes(),
        });
    }
    Ok(output)
}

/// The candidate must read back as the source's paragraphs in `order`, each
/// with its text, paragraph properties and runs, and with every collection
/// the body can anchor unchanged.
fn verify(
    source: &Package,
    candidate: &Package,
    order: &[Slot<'_>],
    removed: Option<&[u8]>,
) -> Result<(), Error> {
    let expected: Vec<_> = source.body().paragraphs().collect();
    let observed: Vec<_> = candidate.body().paragraphs().collect();
    if observed.len() != order.len() {
        return Err(Error::unsupported_source(
            "the spliced paragraph segments do not read back as the expected paragraph count",
        ));
    }
    for (observed, slot) in observed.iter().zip(order) {
        let index = match slot {
            Slot::New(text) => {
                if observed.to_text() != *text {
                    return Err(Error::unsupported_source(
                        "the inserted paragraph does not read back with its text",
                    ));
                }
                continue;
            },
            Slot::Source(index) => index,
        };
        let expected = expected.get(*index).ok_or(Error::unsupported_source(
            "paragraph order is outside the source",
        ))?;
        if observed.to_text() != expected.to_text() {
            return Err(Error::unsupported_source(
                "a spliced paragraph does not read back with its text",
            ));
        }
        if observed.format().raw() != expected.format().raw() {
            return Err(Error::unsupported_source(
                "a paragraph inherits paragraph properties from its neighbours and would change",
            ));
        }
        let mut left = observed.runs();
        let mut right = expected.runs();
        loop {
            match (left.next(), right.next()) {
                (Some(left), Some(right))
                    if left.text() == right.text()
                        && left.format().raw() == right.format().raw() => {},
                (None, None) => break,
                _ => {
                    return Err(Error::unsupported_source(
                        "a paragraph inherits character formatting from its neighbours and would change",
                    ));
                },
            }
        }
    }
    let model = source.model();
    let candidate_model = candidate.model();
    // A removed paragraph takes the content its segment holds whole
    // (fields, pictures, notes, objects are groups inside it), so those
    // collections may shrink, but a range it would split may not, and
    // unmodelled content never goes.
    if let Some(removed) = removed {
        let occurrences = |needle: &[u8]| {
            removed
                .windows(needle.len())
                .filter(|window| *window == needle)
                .count()
        };
        if occurrences(br"\bkmkstart") != occurrences(br"\bkmkend")
            || occurrences(br"\atrfstart") != occurrences(br"\atrfend")
        {
            return Err(Error::unsupported_source(
                "the removed paragraph holds one end of a bookmark or comment range",
            ));
        }
    }
    let kept = |before: usize, after: usize| {
        if removed.is_some() {
            after <= before
        } else {
            after == before
        }
    };
    // Table rows are copied whole (with the state before them replayed
    // when their predecessor changed); they must read back unchanged.
    let same_tables = source.tables().len() == candidate.tables().len()
        && source
            .tables()
            .iter()
            .zip(candidate.tables())
            .all(|(left, right)| format!("{left:?}") == format!("{right:?}"));
    if !same_tables {
        return Err(Error::unsupported_source(
            "a table of the body does not read back unchanged",
        ));
    }
    let same_sections = source.sections().len() == candidate.sections().len()
        && source
            .sections()
            .iter()
            .zip(candidate.sections())
            .all(|(left, right)| format!("{left:?}") == format!("{right:?}"));
    if !same_sections {
        return Err(Error::unsupported_source(
            "a section of the body does not read back unchanged",
        ));
    }
    let same_collections = source.tables().len() == candidate.tables().len()
        && kept(source.fields().len(), candidate.fields().len())
        && kept(source.objects().len(), candidate.objects().len())
        && kept(source.pictures().len(), candidate.pictures().len())
        && kept(source.shapes().len(), candidate.shapes().len())
        && kept(source.revisions().len(), candidate.revisions().len())
        && kept(source.annotations().len(), candidate.annotations().len())
        && kept(source.notes().len(), candidate.notes().len())
        // Content the reader does not model is never dropped with a
        // paragraph.
        && source.opaque().len() == candidate.opaque().len()
        && kept(
            model.bookmarks().bookmarks().len(),
            candidate_model.bookmarks().bookmarks().len(),
        )
        && kept(model.form_fields().len(), candidate_model.form_fields().len())
        && kept(model.math_zones().len(), candidate_model.math_zones().len())
        && model.unknown_syntax_markers() == candidate_model.unknown_syntax_markers();
    if !same_collections {
        return Err(Error::unsupported_source(
            "the paragraph anchors fields, pictures, bookmarks, notes or other content that the splice would drop or split",
        ));
    }
    Ok(())
}
