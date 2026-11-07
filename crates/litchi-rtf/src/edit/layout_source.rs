//! Locate balanced body content without consuming header/trailer destinations.
use super::{Error, Package};
use crate::lexer::{ControlWord, Lexer, Token};
use bumpalo::Bump;
use std::ops::Range;

pub(super) struct Plan {
    pub(super) span: Range<usize>,
    pub(super) event_offsets: Vec<usize>,
}
#[derive(Clone, Copy, PartialEq, Eq)]
enum GroupKind {
    Body,
    Destination,
    Bookmark,
    NestedDestination,
}
struct Group {
    start: usize,
    root_start: Option<usize>,
    kind: GroupKind,
    visible: bool,
}
fn push<T>(items: &mut Vec<T>, value: T) -> Result<(), Error> {
    items
        .try_reserve(1)
        .map_err(|_| Error::Write("could not reserve layout source map".into()))?;
    items.push(value);
    Ok(())
}
fn head<'a, 'text>(tokens: &'a [Token<'text>], open: usize) -> Option<&'a ControlWord<'text>> {
    let mut i = open + 1;
    while matches!(tokens.get(i),Some(Token::Text(s)) if s.trim().is_empty()) {
        i += 1;
    }
    if matches!(
        tokens.get(i),
        Some(Token::Control(ControlWord::IgnorableDestination))
    ) {
        i += 1;
    }
    match tokens.get(i) {
        Some(Token::Control(control)) => Some(control),
        _ => None,
    }
}

pub(super) fn plan(source: &Package, bytes: &[u8]) -> Result<Plan, Error> {
    let input = std::str::from_utf8(bytes)
        .map_err(|_| Error::unsupported_source("layout source is not UTF-8"))?;
    let arena = Bump::new();
    let mut lexer = Lexer::new_with_limits(input, &arena, source.limits())
        .for_document(source.options().leniency_value());
    let (tokens, spans) = lexer.tokenize_with_spans()?;
    // Unknown body destinations are already retained as complete inert groups.
    // Known metadata/story destinations must remain outside a body rewrite.
    let mut opaque = Vec::new();
    opaque
        .try_reserve(source.opaque().len())
        .map_err(|_| Error::Write("could not reserve layout opaque map".into()))?;
    opaque.extend(
        source
            .opaque()
            .iter()
            .filter(|n| {
                n.kind() == crate::opaque::Kind::Destination
                    && matches!(n.anchor(), crate::opaque::Anchor::Body(_))
            })
            .map(|n| n.source()),
    );
    opaque.sort_unstable();
    opaque.dedup();
    let mut groups: Vec<Group> = Vec::new();
    let mut protected = Vec::new();
    let mut event_offsets = Vec::new();
    let mut start = None;
    let mut end = None;
    let mut root_end = None;
    for (index, (token, span)) in tokens.iter().zip(&spans).enumerate() {
        match token {
            Token::OpenBrace => {
                let parent_destination = groups.last().is_some_and(|g| g.kind != GroupKind::Body);
                let bookmark = !parent_destination
                    && matches!(
                        head(&tokens, index),
                        Some(ControlWord::BookmarkStart | ControlWord::BookmarkEnd)
                    );
                let kind = if parent_destination {
                    GroupKind::NestedDestination
                } else if bookmark {
                    GroupKind::Bookmark
                } else if super::producer_paragraphs::opens_destination(
                    bytes
                        .get(span.end..)
                        .ok_or(Error::Internal("layout token span is outside the source"))?,
                ) {
                    GroupKind::Destination
                } else {
                    GroupKind::Body
                };
                if bookmark {
                    push(&mut event_offsets, span.start)?;
                }
                let root_start = if groups.len() == 1 {
                    Some(span.start)
                } else {
                    groups.last().and_then(|g| g.root_start)
                };
                push(
                    &mut groups,
                    Group {
                        start: span.start,
                        root_start,
                        kind,
                        visible: false,
                    },
                )?;
            },
            Token::CloseBrace => {
                let group = groups
                    .pop()
                    .ok_or(Error::unsupported_source("unbalanced layout source group"))?;
                if group.kind == GroupKind::Destination {
                    let range = group.start..span.end;
                    let raw = bytes.get(range.clone()).ok_or(Error::Internal(
                        "layout destination span is outside the source",
                    ))?;
                    if opaque.binary_search(&raw).is_err() {
                        push(&mut protected, range)?;
                    }
                }
                if groups.is_empty() {
                    root_end = Some(span.start);
                }
                if group.visible {
                    if groups.len() == 1 {
                        end = Some(span.end);
                    }
                    if let Some(parent) = groups.last_mut() {
                        parent.visible = true;
                    }
                }
            },
            _ if groups.last().is_some_and(|g| g.kind == GroupKind::Body) => {
                if matches!(
                    token,
                    Token::Control(
                        ControlWord::Page(_)
                            | ControlWord::Column(_)
                            | ControlWord::SoftPageBreak(_)
                            | ControlWord::SoftColumnBreak(_)
                            | ControlWord::SoftLineBreak(_)
                            | ControlWord::SoftLineHeight(_)
                    )
                ) {
                    push(&mut event_offsets, span.start)?;
                }
                let visible = match token {
                    Token::Text(text) => text.bytes().any(|b| !matches!(b, b'\r' | b'\n')),
                    Token::Control(
                        ControlWord::Par
                        | ControlWord::Line
                        | ControlWord::Tab
                        | ControlWord::Unicode(_),
                    ) => true,
                    Token::Control(control) => super::run_control_symbol_text(control).is_some(),
                    Token::Binary(_) => {
                        return Err(Error::unsupported_source(
                            "layout body contains binary data",
                        ));
                    },
                    _ => false,
                };
                if visible {
                    let group = groups
                        .last_mut()
                        .ok_or(Error::Internal("layout body group disappeared"))?;
                    start.get_or_insert(group.root_start.unwrap_or(span.start));
                    end = Some(span.end);
                    group.visible = true;
                }
            },
            _ => {},
        }
    }
    let root_end = root_end.ok_or(Error::unsupported_source(
        "layout source root is not closed",
    ))?;
    let span = match start.zip(end) {
        Some((start, end)) => start..end,
        None if source.text().is_empty() => root_end..root_end,
        None => {
            return Err(Error::unsupported_source(
                "layout source has no body content extent",
            ));
        },
    };
    if protected
        .iter()
        .any(|r| r.start < span.end && r.end > span.start)
    {
        return Err(Error::unsupported_source(
            "paragraph-layout source contains interleaved non-body destinations",
        ));
    }
    if event_offsets.len() != source.model().body_story_events().len() {
        return Err(Error::unsupported_source(
            "layout source events do not match retained body events",
        ));
    }
    Ok(Plan {
        span,
        event_offsets,
    })
}
