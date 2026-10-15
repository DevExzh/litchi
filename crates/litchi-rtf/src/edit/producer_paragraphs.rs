//! Paragraph removal, moves and swaps on producer RTF by exact source
//! splicing (record 0293).
//!
//! The lifecycle projection and the ordinary paragraph map accept only plain
//! ASCII bodies with default formatting, so producer RTF (Word, LibreOffice,
//! WordPad: `\pard` paragraph properties, character-formatting groups,
//! fonts and colours) was refused. This seam instead cuts the body into the
//! source segments its root-level `\par` controls delimit and removes, moves
//! or exchanges whole segments. Every byte outside the moved segments is
//! copied verbatim, so a move and its inverse move restore the source bytes.
//!
//! RTF state is positional: a segment can rely on paragraph or character
//! properties set before it, and anchored content (fields, pictures,
//! bookmarks, tables) can span segments. Rather than model those rules, the
//! candidate is reparsed with the source's open policy and must read back as
//! the expected paragraph list, each paragraph with its exact text, paragraph
//! properties and runs, and with every other story and collection unchanged.
//! Anything else is a typed refusal that names what moved.

use std::ops::Range;

use bumpalo::Bump;

use super::{Commit, Error, Operation, Package, semantic_changes};
use crate::lexer::{ControlWord, Lexer, Token};

/// One paragraph's source segment: from the end of the previous paragraph's
/// boundary through the end of its own `\par` (and the delimiter space that
/// belongs to that control word).
struct Segments {
    ranges: Vec<Range<usize>>,
    /// Whether the last paragraph has no `\par` of its own: its segment ends
    /// before the root group's closing brace.
    unterminated_last: bool,
}

/// Publishes one paragraph removal, move or swap on a producer source.
pub(super) fn commit(source: Package, operation: &Operation) -> Result<Commit, Error> {
    super::ensure_changed_publication_allowed(&source)?;
    let count = source.paragraph_count();
    let (order, touched) = match operation {
        Operation::RemoveParagraph { position, text } => {
            check_text(&source, *position, text)?;
            let order: Vec<usize> = (0..count).filter(|index| index != position).collect();
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
            (order, vec![*position, *final_position])
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
            (order, vec![*first, *second])
        },
        _ => return Err(Error::BodyDestinationConflict),
    };
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::UnsupportedSource(
            "producer paragraph splicing refuses compressed RTF",
        ));
    }
    let input = std::str::from_utf8(source_bytes).map_err(|_error| {
        Error::UnsupportedSource("producer paragraph splicing requires an ASCII or UTF-8 transport")
    })?;
    let cut = segments(&source, input)?;
    if cut.ranges.len() != count {
        return Err(Error::UnsupportedSource(
            "the body's paragraphs are not delimited by root-level \\par controls",
        ));
    }
    let last = count.saturating_sub(1);
    if cut.unterminated_last
        && touched.contains(&last)
        && !matches!(operation, Operation::RemoveParagraph { .. })
    {
        return Err(Error::UnsupportedSource(
            "the last paragraph has no \\par of its own and cannot move",
        ));
    }
    let bytes = splice(source_bytes, &cut.ranges, operation, source.limits())?;
    let snapshot =
        Package::from_bytes_with_options(&bytes, source.options()).map_err(|_error| {
            Error::UnsupportedSource("the spliced paragraph segments do not parse as RTF")
        })?;
    verify(&source, &snapshot, &order)?;
    // The inverse move or swap must be plannable on the candidate the same
    // way, so the edit is reversible through this seam too. (A removal's
    // inverse is an insertion, which this seam does not plan.)
    if matches!(operation, Operation::RemoveParagraph { .. }) {
        let delta = semantic_changes(std::slice::from_ref(operation), &[]);
        return Ok(Commit::new(source, snapshot, 1, delta));
    }
    let candidate_input = std::str::from_utf8(&bytes).map_err(|_error| {
        Error::UnsupportedSource("the spliced paragraph segments are not UTF-8")
    })?;
    if segments(&snapshot, candidate_input)?.ranges.len() != snapshot.paragraph_count() {
        return Err(Error::UnsupportedSource(
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
    let mut first_fallback = None;
    for (token, span) in tokens.iter().zip(&spans) {
        if pending_group_par {
            match token {
                Token::CloseBrace => {
                    depth = depth.saturating_sub(1);
                    destinations.pop();
                    if depth == 1 {
                        boundaries.push(boundary_end(input, span.end));
                        pending_group_par = false;
                    }
                    continue;
                },
                Token::Text(text) if text.trim().is_empty() => continue,
                _ => {
                    return Err(Error::UnsupportedSource(
                        "a paragraph ends inside a group that continues with more content",
                    ));
                },
            }
        }
        let in_destination = destinations.last().copied().unwrap_or(false);
        match token {
            Token::OpenBrace => {
                let destination =
                    in_destination || opens_destination(bytes.get(span.end..).unwrap_or_default());
                if depth == 1 && !destination && first_start.is_none() {
                    first_fallback.get_or_insert(span.start);
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
            },
            Token::Control(ControlWord::Pard) if depth == 1 && boundaries.is_empty() => {
                // The last paragraph reset before the first boundary opens
                // the first paragraph; header destinations precede it.
                first_start = Some(span.start);
            },
            Token::Control(ControlWord::Par) if depth == 1 => {
                if boundaries.is_empty() && first_start.is_none() && first_fallback.is_none() {
                    // An empty first paragraph is its `\par` alone.
                    first_fallback = Some(span.start);
                }
                boundaries.push(boundary_end(input, span.end));
            },
            Token::Text(text) if depth == 1 && first_start.is_none() && !text.trim().is_empty() => {
                // Without a `\pard`, the first paragraph starts at its first
                // text (`{\rtf1 1\par 2\par}`); controls before it stay
                // with the document. Readback refuses a paragraph that needs
                // them.
                first_fallback.get_or_insert(span.start);
            },
            Token::Control(ControlWord::Par) if depth > 1 && !in_destination => {
                pending_group_par = true;
            },
            Token::Control(
                ControlWord::TableRow | ControlWord::TableCell | ControlWord::SectionBreak,
            ) if !in_destination => {
                return Err(Error::UnsupportedSource(
                    "producer paragraph splicing refuses bodies with tables or section breaks",
                ));
            },
            _ => {},
        }
    }
    let first_start = first_start
        .or(first_fallback)
        .ok_or(Error::UnsupportedSource(
            "the first paragraph has no root-level \\pard or text that proves where it starts",
        ))?;
    let root_end = root_end.ok_or(Error::UnsupportedSource("the RTF root group is not closed"))?;
    let mut ranges = Vec::new();
    ranges
        .try_reserve_exact(boundaries.len() + 1)
        .map_err(|_error| Error::Write("could not reserve paragraph segments".to_string()))?;
    let mut start = first_start;
    for end in boundaries {
        if end <= start {
            return Err(Error::UnsupportedSource(
                "a root-level \\par precedes the first paragraph's start",
            ));
        }
        ranges.push(start..end);
        start = end;
    }
    // Text after the last boundary is an unterminated last paragraph.
    let tail = bytes.get(start..root_end).unwrap_or_default();
    let unterminated_last = tail_has_content(tail);
    if unterminated_last {
        ranges.push(start..root_end);
    }
    Ok(Segments {
        ranges,
        unterminated_last,
    })
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
    source.get(range).ok_or(Error::UnsupportedSource(
        "paragraph segment is outside the source",
    ))
}

fn splice(
    source: &[u8],
    ranges: &[Range<usize>],
    operation: &Operation,
    limits: crate::SourceLimits,
) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_error| Error::Write("could not reserve the spliced RTF".to_string()))?;
    let segment = |index: usize| -> Result<&[u8], Error> {
        ranges
            .get(index)
            .and_then(|range| source.get(range.clone()))
            .ok_or(Error::UnsupportedSource(
                "paragraph segment is outside the source",
            ))
    };
    match operation {
        Operation::RemoveParagraph { position, .. } => {
            let range = ranges
                .get(*position)
                .ok_or(Error::UnsupportedSource("paragraph segment is missing"))?;
            output.extend_from_slice(part(source, 0..range.start)?);
            output.extend_from_slice(part(source, range.end..source.len())?);
        },
        Operation::MoveParagraph {
            position,
            final_position,
            ..
        } => {
            let moved = ranges
                .get(*position)
                .ok_or(Error::UnsupportedSource("paragraph segment is missing"))?;
            if final_position > position {
                // Insert after the segment that ends up before it.
                let anchor = ranges
                    .get(*final_position)
                    .ok_or(Error::UnsupportedSource("paragraph segment is missing"))?;
                output.extend_from_slice(part(source, 0..moved.start)?);
                output.extend_from_slice(part(source, moved.end..anchor.end)?);
                output.extend_from_slice(segment(*position)?);
                output.extend_from_slice(part(source, anchor.end..source.len())?);
            } else {
                let anchor = ranges
                    .get(*final_position)
                    .ok_or(Error::UnsupportedSource("paragraph segment is missing"))?;
                output.extend_from_slice(part(source, 0..anchor.start)?);
                output.extend_from_slice(segment(*position)?);
                output.extend_from_slice(part(source, anchor.start..moved.start)?);
                output.extend_from_slice(part(source, moved.end..source.len())?);
            }
        },
        Operation::SwapParagraph { first, second, .. } => {
            let (low, high) = if first < second {
                (*first, *second)
            } else {
                (*second, *first)
            };
            let left = ranges
                .get(low)
                .ok_or(Error::UnsupportedSource("paragraph segment is missing"))?;
            let right = ranges
                .get(high)
                .ok_or(Error::UnsupportedSource("paragraph segment is missing"))?;
            output.extend_from_slice(part(source, 0..left.start)?);
            output.extend_from_slice(segment(high)?);
            output.extend_from_slice(part(source, left.end..right.start)?);
            output.extend_from_slice(segment(low)?);
            output.extend_from_slice(part(source, right.end..source.len())?);
        },
        _ => return Err(Error::BodyDestinationConflict),
    }
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
fn verify(source: &Package, candidate: &Package, order: &[usize]) -> Result<(), Error> {
    let expected: Vec<_> = source.body().paragraphs().collect();
    let observed: Vec<_> = candidate.body().paragraphs().collect();
    if observed.len() != order.len() {
        return Err(Error::UnsupportedSource(
            "the spliced paragraph segments do not read back as the expected paragraph count",
        ));
    }
    for (observed, index) in observed.iter().zip(order) {
        let expected = expected.get(*index).ok_or(Error::UnsupportedSource(
            "paragraph order is outside the source",
        ))?;
        if observed.to_text() != expected.to_text() {
            return Err(Error::UnsupportedSource(
                "a spliced paragraph does not read back with its text",
            ));
        }
        if observed.format().raw() != expected.format().raw() {
            return Err(Error::UnsupportedSource(
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
                    return Err(Error::UnsupportedSource(
                        "a paragraph inherits character formatting from its neighbours and would change",
                    ));
                },
            }
        }
    }
    let model = source.model();
    let candidate_model = candidate.model();
    let same_collections = source.tables().len() == candidate.tables().len()
        && source.fields().len() == candidate.fields().len()
        && source.objects().len() == candidate.objects().len()
        && source.pictures().len() == candidate.pictures().len()
        && source.shapes().len() == candidate.shapes().len()
        && source.revisions().len() == candidate.revisions().len()
        && source.annotations().len() == candidate.annotations().len()
        && source.notes().len() == candidate.notes().len()
        && source.opaque().len() == candidate.opaque().len()
        && model.bookmarks().bookmarks().len() == candidate_model.bookmarks().bookmarks().len()
        && model.form_fields().len() == candidate_model.form_fields().len()
        && model.math_zones().len() == candidate_model.math_zones().len()
        && model.unknown_syntax_markers() == candidate_model.unknown_syntax_markers();
    if !same_collections {
        return Err(Error::UnsupportedSource(
            "the paragraph anchors fields, pictures, bookmarks, notes or other content that the splice would drop or split",
        ));
    }
    Ok(())
}
