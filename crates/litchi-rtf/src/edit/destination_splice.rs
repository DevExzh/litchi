//! Exact-source publication for retained destination text edits.
//!
//! Table-cell, header/footer paragraph, footnote/endnote and comment text
//! edits used to publish through the canonical writer, which reserializes the
//! whole document and emits every table after the body story.  This seam instead
//! splices the replacement text into the retained source at the byte range
//! the parser recorded for the destination story.  Every byte outside the
//! selected stories is copied verbatim.
//!
//! The located range must contain only text-producing tokens at its own
//! group depth and balanced character-formatting groups below it, so the
//! splice cannot change parser state for following content.  The candidate is
//! reparsed with the source's open policy and compared with the source on
//! every retained story before it is published.

use super::{Error, HeaderFooterParagraph, Operation, Package};

use crate::lexer::{Lexer, Token};
use crate::model::provenance::StorySource;
use bumpalo::Bump;
use std::collections::HashMap;
use std::ops::Range;

/// One located source replacement.
struct Splice {
    range: Range<usize>,
    bytes: Vec<u8>,
}

/// Whether every operation in a transaction can use this seam.
pub(super) fn supports(operations: &[Operation]) -> bool {
    !operations.is_empty()
        && operations.iter().all(|operation| {
            matches!(
                operation,
                Operation::TableCellText { path, .. } if path.nested.is_empty()
            ) || matches!(
                operation,
                Operation::HeaderFooterText { .. }
                    | Operation::NoteText { .. }
                    | Operation::AnnotationText { .. }
            )
        })
}

/// Splice every destination replacement into the retained source and return
/// the validated candidate snapshot.
pub(super) fn publish(source: &Package, operations: &[Operation]) -> Result<Package, Error> {
    let source_bytes = source.source_bytes().ok_or(Error::unsupported_operation(
        "snapshot has no exact RTF source",
    ))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::unsupported_source(
            "compressed RTF needs a transport-aware destination rewrite",
        ));
    }
    if !source_bytes.is_ascii() {
        return Err(Error::unsupported_source(
            "exact destination editing requires an ASCII RTF transport",
        ));
    }
    let input = std::str::from_utf8(source_bytes)
        .map_err(|_error| Error::unsupported_source("destination source is not ASCII"))?;
    let arena = Bump::new();
    let mut lexer = Lexer::new_with_limits(input, &arena, source.limits())
        .for_document(source.options().leniency_value());
    let (tokens, spans) = lexer.tokenize_with_spans()?;
    let provenance = source.model().text_provenance();

    let mut splices = Vec::new();
    splices
        .try_reserve_exact(operations.len())
        .map_err(|_error| allocation_error("destination splices", operations.len()))?;
    for operation in operations {
        let (story, after, paragraphs) = match operation {
            Operation::TableCellText { path, after, .. } => {
                if !path.nested.is_empty() {
                    return Err(Error::unsupported_source(
                        "nested table cells have no exact source provenance",
                    ));
                }
                let key = (
                    path.root.table_index,
                    path.root.row_index,
                    path.root.cell_index,
                );
                (provenance.cell(key), after, true)
            },
            Operation::HeaderFooterText { target, after, .. } => (
                header_footer_source(source, *target, provenance),
                after,
                false,
            ),
            Operation::NoteText { index, after, .. } => (provenance.note(*index), after, true),
            Operation::AnnotationText { index, after, .. } => {
                (provenance.annotation(*index), after, true)
            },
            _ => {
                return Err(Error::unsupported_operation(
                    "body splices and retained-destination edits cannot compose safely",
                ));
            },
        };
        let story = story.ok_or(Error::unsupported_source(
            "destination text has no exact source provenance",
        ))?;
        splices.push(locate(
            source_bytes,
            &tokens,
            &spans,
            story,
            after,
            paragraphs,
        )?);
    }
    splices.sort_by_key(|splice| splice.range.start);
    for pair in splices.windows(2) {
        if let [left, right] = pair
            && right.range.start < left.range.end
        {
            return Err(Error::unsupported_source(
                "destination source ranges overlap",
            ));
        }
    }

    let removed = splices
        .iter()
        .map(|splice| splice.range.len())
        .sum::<usize>();
    let added = splices
        .iter()
        .map(|splice| splice.bytes.len())
        .sum::<usize>();
    let candidate_len = source_bytes
        .len()
        .saturating_sub(removed)
        .saturating_add(added);
    let limit = source.limits().max_input_bytes();
    if candidate_len > limit {
        return Err(Error::InputTooLarge {
            observed: candidate_len,
            limit,
        });
    }
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(candidate_len)
        .map_err(|_error| allocation_error("destination candidate", candidate_len))?;
    let mut cursor = 0usize;
    for splice in &splices {
        candidate.extend_from_slice(source_bytes.get(cursor..splice.range.start).ok_or(
            Error::unsupported_source("destination provenance leaves the source"),
        )?);
        candidate.extend_from_slice(&splice.bytes);
        cursor = splice.range.end;
    }
    candidate.extend_from_slice(source_bytes.get(cursor..).ok_or(Error::unsupported_source(
        "destination provenance leaves the source",
    ))?);

    let snapshot = Package::from_bytes_with_options(&candidate, source.options())?;
    verify(source, &snapshot, operations)?;
    Ok(snapshot)
}

fn header_footer_source<'p>(
    source: &Package,
    target: HeaderFooterParagraph,
    provenance: &'p crate::model::provenance::TextProvenance,
) -> Option<&'p StorySource> {
    let story = source
        .sections()
        .get(target.section())?
        .headers_footers
        .iter()
        .position(|candidate| candidate.header_type == target.kind())?;
    provenance.header_footer_paragraph((target.section(), story, target.paragraph()))
}

/// Find the exact source range of one story and render its replacement.
fn locate(
    source: &[u8],
    tokens: &[Token<'_>],
    spans: &[Range<usize>],
    story: &StorySource,
    after: &str,
    paragraphs: bool,
) -> Result<Splice, Error> {
    let range = match (&story.text, story.terminator) {
        (Some(text), _) => text.clone(),
        (None, Some(terminator)) => terminator..terminator,
        (None, None) => {
            return Err(Error::unsupported_source(
                "empty destination text has no insertion boundary",
            ));
        },
    };
    if range.end > source.len() || range.start > range.end {
        return Err(Error::unsupported_source(
            "destination provenance is outside the source",
        ));
    }
    validate_range(source, tokens, spans, &range)?;

    let mut bytes = Vec::new();
    encode_text(after, paragraphs, &mut bytes)?;
    if needs_delimiter(source, tokens, spans, range.start)
        && bytes
            .first()
            .is_some_and(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b' ' | b'-'))
    {
        bytes.insert(0, b' ');
    }
    Ok(Splice { range, bytes })
}

/// Index of the first token whose span ends after `offset`.
fn first_token_at_or_after(spans: &[Range<usize>], offset: usize) -> usize {
    spans.partition_point(|span| span.end <= offset)
}

/// Whether the replacement at `offset` directly follows a control word that
/// carries no delimiter, so a leading letter, digit, hyphen or space would be
/// absorbed into it.
fn needs_delimiter(
    source: &[u8],
    tokens: &[Token<'_>],
    spans: &[Range<usize>],
    offset: usize,
) -> bool {
    let Some(previous) = offset.checked_sub(1) else {
        return false;
    };
    let index = first_token_at_or_after(spans, previous);
    let (Some(Token::Control(_)), Some(span)) = (tokens.get(index), spans.get(index)) else {
        return false;
    };
    if span.start > previous || span.end != offset {
        return false;
    }
    let word = source.get(span.clone()).unwrap_or_default();
    word.get(1).is_some_and(u8::is_ascii_alphabetic)
        && word
            .last()
            .is_some_and(|byte| byte.is_ascii_alphanumeric() || *byte == b'-')
}

/// Accept only text-producing tokens at the range's own depth and balanced
/// character-formatting groups below it.
fn validate_range(
    source: &[u8],
    tokens: &[Token<'_>],
    spans: &[Range<usize>],
    range: &Range<usize>,
) -> Result<(), Error> {
    if range.is_empty() {
        return Ok(());
    }
    let first = first_token_at_or_after(spans, range.start);
    let mut depth = 0usize;
    for (token, span) in tokens.iter().zip(spans).skip(first) {
        if span.start >= range.end {
            break;
        }
        if span.start < range.start || span.end > range.end {
            return Err(Error::unsupported_source(
                "destination provenance splits a source token",
            ));
        }
        let word = source.get(span.clone()).unwrap_or_default();
        match token {
            Token::Text(_) => {},
            Token::OpenBrace => depth = depth.saturating_add(1),
            Token::CloseBrace => {
                depth = depth.checked_sub(1).ok_or(Error::unsupported_source(
                    "destination text closes an enclosing group",
                ))?;
            },
            Token::Binary(_) => {
                return Err(Error::unsupported_source(
                    "destination text contains binary data",
                ));
            },
            Token::Control(_) => {
                let name = control_name(word);
                let admitted = match name {
                    None => true,
                    Some(name) if is_text_control(name) => true,
                    Some(name) => depth > 0 && is_character_format_control(name),
                };
                if word == br"\*" || !admitted {
                    return Err(Error::unsupported_source(
                        "destination text contains a structural or state-changing control",
                    ));
                }
            },
        }
    }
    if depth != 0 {
        return Err(Error::unsupported_source(
            "destination text leaves a group open",
        ));
    }
    Ok(())
}

/// The letters of a control word, or `None` for a control symbol.
fn control_name(word: &[u8]) -> Option<&[u8]> {
    let body = word.strip_prefix(b"\\")?;
    let length = body
        .iter()
        .take_while(|byte| byte.is_ascii_alphabetic())
        .count();
    (length != 0).then(|| body.get(..length).unwrap_or_default())
}

/// Control words that only contribute characters to the current story.
fn is_text_control(name: &[u8]) -> bool {
    matches!(
        name,
        b"u" | b"tab"
            | b"line"
            | b"par"
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
    )
}

/// Character-formatting controls whose effect ends with their group.
fn is_character_format_control(name: &[u8]) -> bool {
    matches!(
        name,
        b"b" | b"i"
            | b"ul"
            | b"uld"
            | b"uldash"
            | b"uldashd"
            | b"uldashdd"
            | b"uldb"
            | b"ulnone"
            | b"ulth"
            | b"ulw"
            | b"ulwave"
            | b"strike"
            | b"striked"
            | b"scaps"
            | b"caps"
            | b"v"
            | b"outl"
            | b"shad"
            | b"embo"
            | b"impr"
            | b"sub"
            | b"super"
            | b"nosupersub"
            | b"up"
            | b"dn"
            | b"f"
            | b"fcs"
            | b"fs"
            | b"cf"
            | b"cb"
            | b"chcbpat"
            | b"highlight"
            | b"plain"
            | b"lang"
            | b"langfe"
            | b"langnp"
            | b"langfenp"
            | b"noproof"
            | b"expnd"
            | b"expndtw"
            | b"kerning"
            | b"charscalex"
            | b"cs"
            | b"rtlch"
            | b"ltrch"
            | b"loch"
            | b"hich"
            | b"dbch"
            | b"af"
            | b"afs"
            | b"ab"
            | b"ai"
            | b"uc"
            | b"insrsid"
            | b"charrsid"
            | b"delrsid"
    )
}

/// Render replacement text as RTF transport bytes.
pub(super) fn encode_text(text: &str, paragraphs: bool, output: &mut Vec<u8>) -> Result<(), Error> {
    output
        .try_reserve(text.len().saturating_add(8))
        .map_err(|_error| allocation_error("destination replacement", text.len()))?;
    for character in text.chars() {
        match character {
            '\\' => output.extend_from_slice(br"\\"),
            '{' => output.extend_from_slice(br"\{"),
            '}' => output.extend_from_slice(br"\}"),
            '\t' => output.extend_from_slice(br"\tab "),
            '\n' if paragraphs => output.extend_from_slice(br"\par "),
            '\n' => {
                return Err(Error::unsupported_source(
                    "this destination cannot contain a paragraph break",
                ));
            },
            ' '..='~' => {
                let mut buffer = [0u8; 4];
                output.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
            character if character.is_control() => {
                return Err(Error::unsupported_source(
                    "destination text contains an unsupported control character",
                ));
            },
            character => {
                // A scoped skip count keeps the fallback independent of the
                // surrounding `\ucN` state.
                output.extend_from_slice(br"{\uc1");
                let mut units = [0u16; 2];
                for unit in character.encode_utf16(&mut units) {
                    output.extend_from_slice(br"\u");
                    #[allow(
                        clippy::cast_possible_wrap,
                        reason = "RTF \\uN parameters are signed 16-bit UTF-16 units"
                    )]
                    let signed = *unit as i16;
                    output.extend_from_slice(signed.to_string().as_bytes());
                    output.push(b'?');
                }
                output.push(b'}');
            },
        }
    }
    Ok(())
}

/// The replacement text of every target, keyed for constant-time lookups.
#[derive(Default)]
struct Overrides<'a> {
    cells: HashMap<(usize, usize, usize), &'a str>,
    stories: HashMap<(usize, usize, usize), &'a str>,
    notes: HashMap<usize, &'a str>,
    comments: HashMap<usize, &'a str>,
}

impl<'a> Overrides<'a> {
    fn of(source: &Package, operations: &'a [Operation]) -> Self {
        let mut overrides = Self::default();
        for operation in operations {
            match operation {
                Operation::TableCellText { path, after, .. } if path.nested.is_empty() => {
                    overrides.cells.insert(
                        (
                            path.root.table_index,
                            path.root.row_index,
                            path.root.cell_index,
                        ),
                        after.as_str(),
                    );
                },
                Operation::HeaderFooterText { target, after, .. } => {
                    if let Some(story) =
                        source.sections().get(target.section()).and_then(|section| {
                            section
                                .headers_footers
                                .iter()
                                .position(|story| story.header_type == target.kind())
                        })
                    {
                        overrides.stories.insert(
                            (target.section(), story, target.paragraph()),
                            after.as_str(),
                        );
                    }
                },
                Operation::NoteText { index, after, .. } => {
                    overrides.notes.insert(*index, after.as_str());
                },
                Operation::AnnotationText { index, after, .. } => {
                    overrides.comments.insert(*index, after.as_str());
                },
                _ => {},
            }
        }
        overrides
    }
}

/// Compare the reparsed candidate with the source on every retained story.
fn verify(source: &Package, candidate: &Package, operations: &[Operation]) -> Result<(), Error> {
    let changed = |reason: &'static str| Error::unsupported_source(reason);
    let overrides = Overrides::of(source, operations);
    if source.text() != candidate.text() || source.paragraph_count() != candidate.paragraph_count()
    {
        return Err(changed("destination edit changed the body story"));
    }

    let source_tables = source.tables();
    let candidate_tables = candidate.tables();
    if source_tables.len() != candidate_tables.len() {
        return Err(changed("destination edit changed the table structure"));
    }
    for (table_index, (left, right)) in source_tables.iter().zip(candidate_tables).enumerate() {
        if left.row_count() != right.row_count() {
            return Err(changed("destination edit changed the table structure"));
        }
        for (row_index, (left, right)) in left.rows().iter().zip(right.rows()).enumerate() {
            if left.cell_count() != right.cell_count() {
                return Err(changed("destination edit changed the table structure"));
            }
            for (cell_index, (left, right)) in left.cells().iter().zip(right.cells()).enumerate() {
                let expected = overrides
                    .cells
                    .get(&(table_index, row_index, cell_index))
                    .copied()
                    .unwrap_or_else(|| left.text());
                if right.text() != expected
                    || left.nested_tables().len() != right.nested_tables().len()
                    || left.right_boundary() != right.right_boundary()
                    || left.merge() != right.merge()
                {
                    return Err(changed(
                        "destination edit changed table-cell content outside its target",
                    ));
                }
            }
        }
    }

    let source_sections = source.sections();
    let candidate_sections = candidate.sections();
    if source_sections.len() != candidate_sections.len() {
        return Err(changed("destination edit changed the section structure"));
    }
    for (section_index, (left, right)) in source_sections.iter().zip(candidate_sections).enumerate()
    {
        if left.headers_footers.len() != right.headers_footers.len() {
            return Err(changed(
                "destination edit changed the header/footer structure",
            ));
        }
        for (story_index, (left, right)) in left
            .headers_footers
            .iter()
            .zip(&right.headers_footers)
            .enumerate()
        {
            if left.header_type != right.header_type
                || left.paragraphs.len() != right.paragraphs.len()
            {
                return Err(changed(
                    "destination edit changed the header/footer structure",
                ));
            }
            for (paragraph_index, (left_paragraph, right_paragraph)) in
                left.paragraphs.iter().zip(&right.paragraphs).enumerate()
            {
                let expected = overrides
                    .stories
                    .get(&(section_index, story_index, paragraph_index))
                    .copied()
                    .unwrap_or_else(|| left_paragraph.text.as_ref());
                if right_paragraph.text != expected
                    || left_paragraph.paragraph != right_paragraph.paragraph
                {
                    return Err(changed(
                        "destination edit changed header/footer content outside its target",
                    ));
                }
            }
        }
    }

    let source_notes = source.notes();
    let candidate_notes = candidate.notes();
    if source_notes.len() != candidate_notes.len() {
        return Err(changed("destination edit changed the note structure"));
    }
    for (index, (left, right)) in source_notes.iter().zip(candidate_notes).enumerate() {
        let expected = overrides
            .notes
            .get(&index)
            .copied()
            .unwrap_or_else(|| left.content.as_ref());
        if right.content != expected
            || left.is_footnote != right.is_footnote
            || left.reference != right.reference
            || left.position != right.position
        {
            return Err(changed(
                "destination edit changed note content outside its target",
            ));
        }
    }

    let annotations_equal = source.annotations().len() == candidate.annotations().len()
        && source
            .annotations()
            .iter()
            .zip(candidate.annotations())
            .enumerate()
            .all(|(index, (left, right))| {
                let expected = overrides
                    .comments
                    .get(&index)
                    .copied()
                    .unwrap_or_else(|| left.text.as_ref());
                right.text == expected
                    && left.position == right.position
                    && left.author == right.author
                    && left.id == right.id
            });
    let fields_equal = source.fields().len() == candidate.fields().len()
        && source
            .fields()
            .iter()
            .zip(candidate.fields())
            .all(|(left, right)| {
                left.instruction == right.instruction
                    && left.result == right.result
                    && left.position == right.position
            });
    if !annotations_equal
        || !fields_equal
        || source.pictures() != candidate.pictures()
        || source.shapes().len() != candidate.shapes().len()
        || source.objects().len() != candidate.objects().len()
        || source.styles().len() != candidate.styles().len()
        || source.lists().len() != candidate.lists().len()
        || source.opaque().len() != candidate.opaque().len()
        || source.revisions().len() != candidate.revisions().len()
    {
        return Err(changed(
            "destination edit changed unrelated document content",
        ));
    }
    Ok(())
}

fn allocation_error(resource: &'static str, requested: usize) -> Error {
    Error::Rtf(crate::Error::AllocationFailed {
        resource,
        requested,
    })
}
