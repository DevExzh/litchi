//! The match plan: which pieces change, to what, and which bytes to format.

use std::borrow::Cow;
use std::ops::Range;

use super::Matcher;
#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    test
))]
use crate::error::Code;
use crate::error::{Error, Result};

/// One searchable text piece: the text of one run (one leaf text node).
///
/// Pieces with the same `group` are adjacent text that a match may span;
/// a new group starts after anything a match must not cross. Groups are
/// contiguous and ascending in the slice handed to [`plan`].
#[derive(Clone, Copy)]
pub(crate) struct Piece<'t> {
    /// The piece's current text.
    pub(crate) text: &'t str,
    /// The group the piece belongs to.
    pub(crate) group: usize,
    /// Whether the piece is inside the find's scope. A match is kept only
    /// when every piece it touches is in scope; the pieces outside still
    /// give a whole-word check its context (OfficeCLI's run scope).
    pub(crate) scoped: bool,
}

impl<'t> Piece<'t> {
    /// A piece of `group` inside the scope.
    #[cfg(test)]
    pub(crate) const fn new(text: &'t str, group: usize) -> Self {
        Self {
            text,
            group,
            scoped: true,
        }
    }
}

/// The change to one piece.
#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    test
))]
#[cfg_attr(
    not(any(feature = "pptx", test)),
    allow(
        dead_code,
        reason = "only PPTX formats matched ranges; the other adapters read the text"
    )
)]
pub(crate) struct Rewrite {
    /// Index of the piece in the slice handed to [`plan`].
    pub(crate) piece: usize,
    /// The piece's new text; `None` when only formatting changes.
    pub(crate) text: Option<String>,
    /// Byte ranges of the new text that the format properties cover, in
    /// ascending order. Empty without format properties.
    pub(crate) marks: Vec<Range<usize>>,
}

/// The result of [`plan`].
#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    test
))]
#[derive(Default)]
pub(crate) struct Plan {
    /// Number of matches kept.
    pub(crate) matches: usize,
    /// The changed pieces, in piece order.
    pub(crate) rewrites: Vec<Rewrite>,
}

/// Plan a find over `pieces`.
///
/// `replacement` is `None` for a format-only find; `marks` asks for the
/// byte ranges a format property covers. `max_bytes` bounds the total size
/// of the rewritten texts.
///
/// # Errors
///
/// `limit_exceeded` when the rewritten texts exceed `max_bytes`, and
/// `internal_error` on allocation failure.
#[cfg(test)]
pub(crate) fn plan(
    pieces: &[Piece<'_>],
    matcher: &Matcher<'_>,
    replacement: Option<&str>,
    marks: bool,
    max_bytes: usize,
) -> Result<Plan> {
    plan_with(
        &mut Scratch::default(),
        pieces,
        matcher,
        replacement,
        marks,
        max_bytes,
    )
}

/// Buffers [`plan_with`] reuses across calls, so an adapter planning one
/// paragraph at a time allocates them once.
#[derive(Default)]
pub(crate) struct Scratch {
    joined: String,
    starts: Vec<usize>,
    found: Vec<Range<usize>>,
}

/// [`plan`] with caller-owned buffers.
///
/// # Errors
///
/// As [`plan`].
#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    test
))]
pub(crate) fn plan_with(
    scratch: &mut Scratch,
    pieces: &[Piece<'_>],
    matcher: &Matcher<'_>,
    replacement: Option<&str>,
    marks: bool,
    max_bytes: usize,
) -> Result<Plan> {
    let mut result = Plan::default();
    let mut staged = 0usize;
    let values = replacement.map_or(Values::Format, Values::Same);
    matches_with(
        scratch,
        pieces,
        matcher,
        |members, starts, text, found, first| {
            result.matches = result.matches.saturating_add(found.len());
            rewrite_group(
                members,
                starts,
                text,
                found,
                values,
                marks,
                first,
                &mut result.rewrites,
                &mut staged,
                max_bytes,
            )
        },
    )?;
    Ok(result)
}

/// Find the matches of every group of `pieces`, calling `each` for a
/// group with at least one: its pieces, their offsets in its joined text,
/// that text, the matches (ranges of it) and the index of its first piece.
/// A match is kept only when every non-empty piece it touches is scoped.
///
/// # Errors
///
/// The errors of `each`, and `internal_error` on allocation failure.
pub(crate) fn matches_with(
    scratch: &mut Scratch,
    pieces: &[Piece<'_>],
    matcher: &Matcher<'_>,
    mut each: impl FnMut(&[Piece<'_>], &[usize], &str, &[Range<usize>], usize) -> Result<()>,
) -> Result<()> {
    let Scratch {
        joined,
        starts,
        found,
    } = scratch;
    let mut first = 0;
    while first < pieces.len() {
        let group = pieces[first].group;
        let end = pieces[first..]
            .iter()
            .position(|piece| piece.group != group)
            .map_or(pieces.len(), |offset| first + offset);
        let members = &pieces[first..end];
        // A group of one piece is searched in place.
        let text: Cow<'_, str> = if let [only] = members {
            starts.clear();
            starts.push(0);
            Cow::Borrowed(only.text)
        } else {
            joined.clear();
            starts.clear();
            for piece in members {
                starts.push(joined.len());
                joined.push_str(piece.text);
            }
            Cow::Borrowed(joined.as_str())
        };
        found.clear();
        for range in matcher.find_iter(&text) {
            if members.iter().zip(starts.iter()).all(|(piece, &start)| {
                let piece_end = start + piece.text.len();
                piece.scoped
                    || piece.text.is_empty()
                    || piece_end <= range.start
                    || start >= range.end
            }) {
                found.try_reserve(1).map_err(allocation)?;
                found.push(range);
            }
        }
        if !found.is_empty() {
            each(members, starts, &text, found, first)?;
        }
        first = end;
    }
    Ok(())
}

/// The new text of the ranges a rewrite changes.
#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    test
))]
#[derive(Clone, Copy)]
pub(crate) enum Values<'v> {
    /// Keep the matched text: only formatting changes.
    Format,
    /// One replacement for every range.
    Same(&'v str),
    /// One replacement per range, in order.
    #[cfg_attr(
        not(any(
            feature = "pptx",
            feature = "xlsx",
            feature = "rtf",
            feature = "doc",
            feature = "ppt",
            feature = "xls",
            feature = "xlsb",
            feature = "odt",
            feature = "ods",
            feature = "odp"
        )),
        allow(dead_code, reason = "template merge plans runs through it")
    )]
    Each(&'v [&'v str]),
}

#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    test
))]
impl<'v> Values<'v> {
    fn at(self, index: usize) -> Option<&'v str> {
        match self {
            Self::Format => None,
            Self::Same(text) => Some(text),
            Self::Each(texts) => texts.get(index).copied(),
        }
    }
}

/// Rewrite `pieces` (contiguous text: each piece follows the previous one)
/// for `ranges` of their joined text, each replaced by its value: a range
/// collapses into the piece where it starts, as a find does. Template merge
/// plans its placeholders through this (record 0285).
///
/// # Errors
///
/// `limit_exceeded` beyond `max_bytes` of rewritten text.
#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
pub(crate) fn rewrite_ranges(
    pieces: &[Piece<'_>],
    ranges: &[Range<usize>],
    values: Values<'_>,
    max_bytes: usize,
) -> Result<Vec<Rewrite>> {
    let mut starts = Vec::with_capacity(pieces.len());
    let mut text = String::new();
    for piece in pieces {
        starts.push(text.len());
        text.push_str(piece.text);
    }
    let mut rewrites = Vec::new();
    let mut staged = 0usize;
    rewrite_group(
        pieces,
        &starts,
        &text,
        ranges,
        values,
        false,
        0,
        &mut rewrites,
        &mut staged,
        max_bytes,
    )?;
    Ok(rewrites)
}

#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    test
))]
#[allow(
    clippy::too_many_arguments,
    reason = "one group's inputs and the shared accumulators"
)]
fn rewrite_group(
    members: &[Piece<'_>],
    starts: &[usize],
    text: &str,
    found: &[Range<usize>],
    values: Values<'_>,
    marks: bool,
    offset: usize,
    rewrites: &mut Vec<Rewrite>,
    staged: &mut usize,
    max_bytes: usize,
) -> Result<()> {
    let mut next = 0;
    for (index, (piece, &start)) in members.iter().zip(starts).enumerate() {
        let end = start + piece.text.len();
        // Skip matches that end before this piece.
        while next < found.len() && found[next].end <= start {
            next += 1;
        }
        if start == end || found.get(next).is_none_or(|range| range.start >= end) {
            continue;
        }
        let mut output = String::new();
        let mut covered = Vec::new();
        let mut cursor = start;
        for (at, range) in found
            .iter()
            .enumerate()
            .skip(next)
            .take_while(|(_, range)| range.start < end)
        {
            let from = range.start.max(start);
            output.push_str(&text[cursor..from]);
            match values.at(at) {
                Some(replacement) => {
                    if range.start >= start {
                        if marks && !replacement.is_empty() {
                            covered.push(output.len()..output.len() + replacement.len());
                        }
                        output.push_str(replacement);
                    }
                },
                None => {
                    let to = range.end.min(end);
                    if marks {
                        covered.push(output.len()..output.len() + (to - from));
                    }
                    output.push_str(&text[from..to]);
                },
            }
            cursor = range.end.min(end);
        }
        output.push_str(&text[cursor..end]);
        let changed = output != piece.text;
        if !changed && covered.is_empty() {
            continue;
        }
        *staged = staged.saturating_add(output.len());
        if *staged > max_bytes {
            return Err(Error::new(
                Code::LimitExceeded,
                format!("find/replace would stage more than {max_bytes} bytes of text"),
            ));
        }
        rewrites.try_reserve(1).map_err(allocation)?;
        rewrites.push(Rewrite {
            piece: offset + index,
            text: changed.then_some(output),
            marks: covered,
        });
    }
    Ok(())
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::from(crate::transaction::Error::Allocation { source })
}
