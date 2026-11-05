//! Errors of path resolution (record 0010 §4.4).
//!
//! Every function here runs on the error path only: it may allocate, format
//! and compute suggestions. Messages identify the problem and byte offset;
//! suggestions use static vocabulary or structural counts, never authored
//! paths, identities, sheet names or document text.
//!
//! | Condition | Code |
//! |---|---|
//! | A step name no vocabulary knows | `invalid_path`, with "did you mean" by edit distance |
//! | A name another family knows (`/slide[1]` on a DOCX) | `unsupported_path`, with the family's root forms |
//! | A name the format lacks as a concept (`ABSENT`) | `unsupported_path` |
//! | A name the format has but the adapter does not model, or a gap | `unsupported_capability`, with the follow-up |
//! | A kind that cannot occur under its parent | `unsupported_path` |
//! | Nothing at that position, identity or name | `path_not_found`, with what exists |
//! | A bad `@key`, a content filter, stacked or misplaced predicates, a coordinate outside the grid, an ambiguous identity | `invalid_path` |

use std::fmt::{self, Write as _};

use litchi_core::FileFormat;
use litchi_core::selector::{CellRange, Path, Step};

use super::arg::{Origin, PathArg};
use super::navigate::{Available, ElementKind, GapInfo, Lookup, Navigate, Resolved};
use super::vocabulary::{ElementRule, Vocabulary};
use crate::{Code, Error, Family};

/// Where resolution failed: the path, the failing step and the format.
#[derive(Clone, Copy)]
pub(crate) struct Spot<'a, 'p> {
    arg: &'a PathArg<'p>,
    step: usize,
    format: FileFormat,
}

impl<'a, 'p> Spot<'a, 'p> {
    pub(crate) fn new(arg: &'a PathArg<'p>, step: usize, format: FileFormat) -> Self {
        Self { arg, step, format }
    }

    /// The byte offset of the failing step in the text the path came from,
    /// or in its canonical spelling for a parsed `Path`.
    pub(crate) fn offset(&self) -> usize {
        match self.arg.origin() {
            Origin::Typed => canonical_offset(self.arg.path(), self.step),
            Origin::Text(text) => {
                let first = usize::from(self.arg.is_absolute());
                step_start(text, first, self.step)
            },
            Origin::Sheet { text, rest } => match self.step {
                0 => 0,
                step => step_start(text, rest, step - 1),
            },
        }
    }

    /// `at byte N of "<path>"`.
    fn context(&self) -> String {
        format!("at byte {}", self.offset())
    }

    /// The canonical spelling of the steps before the failing one: `/body`,
    /// `/` for the root, `the start node` for an empty relative prefix.
    fn prefix(&self) -> String {
        "the parent element".to_owned()
    }

    /// The failing step as written in canonical form: `p[99]`.
    fn step(&self) -> String {
        "the requested step".to_owned()
    }

    fn error(&self, code: Code, problem: fmt::Arguments<'_>) -> Error {
        Error::new(code, format!("{problem} {}", self.context())).with_format(self.format)
    }
}

/// Counts bytes written, so canonical offsets need no buffer.
struct Count(usize);

impl fmt::Write for Count {
    fn write_str(&mut self, text: &str) -> fmt::Result {
        self.0 += text.len();
        Ok(())
    }
}

/// The offset of step `step` in the canonical spelling of `path`.
fn canonical_offset(path: &Path<'_>, step: usize) -> usize {
    let mut count = Count(usize::from(path.is_absolute()));
    for before in path.steps().iter().take(step) {
        let _ = write!(count, "{before}/");
    }
    count.0
}

/// The byte offset where step `index` starts in `text`, counting steps from
/// byte `first`. Quoted names and bracketed predicates may contain `/`; the
/// scan skips them with the grammar's quoting rules (record 0005).
fn step_start(text: &str, first: usize, index: usize) -> usize {
    let bytes = text.as_bytes();
    let mut at = first.min(bytes.len());
    for _ in 0..index {
        if at < bytes.len() && is_quote(bytes[at]) {
            at = skip_quoted(bytes, at);
        }
        let mut depth = 0_usize;
        while at < bytes.len() {
            match bytes[at] {
                b'[' => depth += 1,
                b']' => depth = depth.saturating_sub(1),
                quote if depth > 0 && is_quote(quote) => {
                    at = skip_quoted(bytes, at);
                    continue;
                },
                b'/' if depth == 0 => break,
                _ => {},
            }
            at += 1;
        }
        if at >= bytes.len() {
            return bytes.len();
        }
        at += 1;
    }
    at
}

const fn is_quote(byte: u8) -> bool {
    byte == b'"' || byte == b'\''
}

/// The offset just past the quoted span that opens at `open`.
fn skip_quoted(bytes: &[u8], open: usize) -> usize {
    let quote = bytes[open];
    let mut at = open + 1;
    while at < bytes.len() {
        match bytes[at] {
            b'\\' => at += 2,
            byte if byte == quote => return at + 1,
            _ => at += 1,
        }
    }
    bytes.len()
}

/// The last worksheet row and column (Excel 2007 and later).
pub(crate) const MAX_ROW: usize = 1_048_576;
pub(crate) const MAX_COLUMN: usize = 16_384;

fn family_noun(family: Family) -> &'static str {
    match family {
        Family::Document => "word-processing documents",
        Family::Presentation => "presentations",
        Family::Workbook => "workbooks",
        _ => "other documents",
    }
}

fn roots_suggestion(spot: &Spot<'_, '_>, vocabulary: &Vocabulary) -> String {
    format!("{:?} paths start at {}", spot.format, vocabulary.roots)
}

/// A document-level path that does not start with `/`.
pub(crate) fn relative(arg: &PathArg<'_>, format: FileFormat) -> Error {
    let spot = Spot::new(arg, 0, format);
    spot.error(
        Code::InvalidPath,
        format_args!("a document path must start with '/'"),
    )
    .with_suggestion("prefix the path with '/'")
}

/// A step name no vocabulary knows.
pub(crate) fn unknown(spot: &Spot<'_, '_>, name: &str, vocabulary: &Vocabulary) -> Error {
    if vocabulary.family == Family::Workbook && looks_like_cell(name) {
        return spot
            .error(
                Code::InvalidPath,
                format_args!("cell reference is outside the worksheet grid A1:XFD1048576"),
            )
            .with_suggestion("columns run from A to XFD and rows from 1 to 1048576");
    }
    let error = spot.error(Code::InvalidPath, format_args!("unknown element"));
    error.with_suggestion(did_you_mean(name, vocabulary))
}

/// Letters then digits, optionally twice around `:`: an A1 reference that
/// failed `CellRange::parse`, so it lies outside the grid.
fn looks_like_cell(name: &str) -> bool {
    fn reference(text: &str) -> bool {
        let letters = text.bytes().take_while(u8::is_ascii_alphabetic).count();
        letters > 0
            && letters < text.len()
            && text.bytes().skip(letters).all(|byte| byte.is_ascii_digit())
    }
    match name.split_once(':') {
        Some((start, end)) => reference(start) && reference(end),
        None => reference(name),
    }
}

/// A name another family knows.
pub(crate) fn foreign(
    spot: &Spot<'_, '_>,
    _name: &str,
    own: &Vocabulary,
    other: &Vocabulary,
) -> Error {
    spot.error(
        Code::UnsupportedPath,
        format_args!(
            "{:?} documents do not have this element, which belongs to {}",
            spot.format,
            family_noun(other.family)
        ),
    )
    .with_suggestion(roots_suggestion(spot, own))
}

/// Worksheet or cell addressing on a format that has no worksheets.
pub(crate) fn no_sheets(spot: &Spot<'_, '_>, own: &Vocabulary) -> Error {
    spot.error(
        Code::UnsupportedPath,
        format_args!("{:?} documents have no worksheets or cells", spot.format),
    )
    .with_suggestion(roots_suggestion(spot, own))
}

/// A name the format lacks as a concept.
pub(crate) fn absent(spot: &Spot<'_, '_>, canonical: &str, own: &Vocabulary) -> Error {
    spot.error(
        Code::UnsupportedPath,
        format_args!(
            "{:?} documents have no \"{canonical}\" elements",
            spot.format
        ),
    )
    .with_suggestion(roots_suggestion(spot, own))
}

/// A name the format has but its adapter does not model.
pub(crate) fn unmodeled(spot: &Spot<'_, '_>, canonical: &str) -> Error {
    spot.error(
        Code::UnsupportedCapability,
        format_args!(
            "reading \"{canonical}\" elements from {:?} is not supported yet",
            spot.format
        ),
    )
    .with_suggestion("capabilities() lists the elements Litchi reads from this format")
}

/// A predicate on a step that names a child by free text.
pub(crate) fn named_with_predicates(spot: &Spot<'_, '_>, child: &str) -> Error {
    spot.error(
        Code::InvalidPath,
        format_args!("a {child} is named by the step alone, without predicates"),
    )
    .with_suggestion("write the child name alone")
}

/// An `@key` the element does not accept as an identity.
pub(crate) fn not_identity(spot: &Spot<'_, '_>, rule: &ElementRule, key: &str) -> Error {
    let _ = key;
    spot.error(
        Code::InvalidPath,
        format_args!("attribute is not an identity of {}", rule.name),
    )
    .with_suggestion(identity_forms(rule))
}

/// `[@key]` without a value.
pub(crate) fn identity_without_value(spot: &Spot<'_, '_>, rule: &ElementRule, key: &str) -> Error {
    let _ = key;
    spot.error(Code::InvalidPath, format_args!("identity needs a value"))
        .with_suggestion(identity_forms(rule))
}

/// A content filter in a path.
pub(crate) fn content_filter(spot: &Spot<'_, '_>) -> Error {
    spot.error(
        Code::InvalidPath,
        format_args!("{} filters by content", spot.step()),
    )
    .with_suggestion(
        "get and find address one element by position or identity; select by content with query",
    )
}

/// Several predicates on one step that are not all identities.
pub(crate) fn stacked(spot: &Spot<'_, '_>) -> Error {
    spot.error(
        Code::InvalidPath,
        format_args!("{} stacks predicates that are not identities", spot.step()),
    )
    .with_suggestion("use one [N], [last()] or bare token, or stack only [@key=value] identities")
}

/// A bare token the element does not accept.
pub(crate) fn bare_token(spot: &Spot<'_, '_>, rule: &ElementRule, token: &str) -> Error {
    let _ = token;
    spot.error(
        Code::InvalidPath,
        format_args!("{} does not take this bare token", rule.name),
    )
    .with_suggestion(identity_forms(rule))
}

/// Column letters outside `A..XFD`.
pub(crate) fn bad_column(spot: &Spot<'_, '_>, token: &str) -> Error {
    let _ = token;
    spot.error(Code::InvalidPath, format_args!("invalid worksheet column"))
        .with_suggestion("columns run from A to XFD, e.g. col[B] or col[2]")
}

/// A row or column number outside the grid.
pub(crate) fn out_of_grid(spot: &Spot<'_, '_>, axis: &str, number: usize, max: usize) -> Error {
    spot.error(
        Code::InvalidPath,
        format_args!("{axis} {number} is outside the worksheet grid"),
    )
    .with_suggestion(format!("{axis}s run from 1 to {max}"))
}

/// An index with no 1-based spelling (only from a hand-built `Path`).
pub(crate) fn index_too_large(spot: &Spot<'_, '_>) -> Error {
    spot.error(Code::InvalidPath, format_args!("the index is too large"))
}

/// An adapter kind whose name is not in its vocabulary: a bug in the
/// adapter, reported instead of panicking.
pub(crate) fn kind_without_rule(spot: &Spot<'_, '_>, name: &str) -> Error {
    spot.error(
        Code::InternalError,
        format_args!("element kind \"{name}\" is missing from its vocabulary"),
    )
}

/// The forms an element accepts, for suggestions: `p[N]`, `p[last()]`,
/// `p[@paraId=...]`, ...
fn identity_forms(rule: &ElementRule) -> String {
    let kind = rule.name;
    let mut forms = format!("{kind} accepts {kind}[N], {kind}[last()]");
    for identity in rule
        .identities
        .iter()
        .filter(|identity| !identity.is_native())
    {
        let _ = write!(forms, ", {kind}[@{}=...]", identity.key);
    }
    forms
}

/// The error for what [`Navigate::child`] answered other than found or
/// missing.
pub(crate) fn lookup<'f, N: Navigate<'f>>(
    nav: &N,
    spot: &Spot<'_, '_>,
    parent: &N::Node,
    wanted: &Resolved<'_, N::Kind>,
    answer: Lookup<N::Node, N::Gap>,
) -> Error {
    let kind = wanted.kind.name();
    match answer {
        Lookup::NotApplicable => {
            let error = spot.error(
                Code::UnsupportedPath,
                format_args!("{kind} elements cannot occur under {}", spot.prefix()),
            );
            with_available(nav, spot, parent, wanted, error)
        },
        Lookup::Gap(gap) => spot
            .error(
                Code::UnsupportedCapability,
                format_args!(
                    "reading \"{kind}\" elements from {:?} is not supported yet",
                    spot.format
                ),
            )
            .with_suggestion(format!(
                "follow-up {}; capabilities() lists this gap",
                gap.follow_up()
            )),
        Lookup::Ambiguous(count) => {
            let error = spot.error(
                Code::InvalidPath,
                format_args!(
                    "{} matches {count} elements under {}",
                    spot.step(),
                    spot.prefix()
                ),
            );
            let mut available = Available::default();
            let listed =
                nav.available(parent, wanted, &mut available).is_ok() && !available.is_empty();
            let mut suggestion = String::from("add another identity predicate to pick one");
            if listed {
                let _ = write!(
                    suggestion,
                    "; available at {}: {}",
                    spot.prefix(),
                    available.render()
                );
            }
            error.with_suggestion(suggestion)
        },
        Lookup::Found(_) | Lookup::Missing => spot.error(
            Code::InternalError,
            format_args!("the resolver mishandled a lookup of {kind}"),
        ),
    }
}

/// The `path_not_found` error of a miss.
pub(crate) fn not_found<'f, N: Navigate<'f>>(
    nav: &N,
    spot: &Spot<'_, '_>,
    parent: &N::Node,
    wanted: &Resolved<'_, N::Kind>,
) -> Error {
    let error = spot.error(
        Code::PathNotFound,
        format_args!("requested {} was not found", wanted.kind.name()),
    );
    with_available(nav, spot, parent, wanted, error)
}

fn with_available<'f, N: Navigate<'f>>(
    nav: &N,
    spot: &Spot<'_, '_>,
    parent: &N::Node,
    wanted: &Resolved<'_, N::Kind>,
    error: Error,
) -> Error {
    let mut available = Available::default();
    // A failure while describing the parent leaves the suggestion out; the
    // error being reported is the miss.
    if nav.available(parent, wanted, &mut available).is_err() || available.is_empty() {
        return error;
    }
    let suggestion = format!("available at {}: {}", spot.prefix(), available.render());
    error.with_suggestion(suggestion)
}

/// Whether a step at or after `from` is an A1 cell or range: `/Sheet1/A1`
/// on a format without worksheets.
pub(crate) fn addresses_cells(steps: &[Step<'_>], from: usize) -> bool {
    steps
        .iter()
        .skip(from)
        .any(|step| step.predicates().is_empty() && CellRange::parse(step.name()).is_some())
}

/// The "did you mean" suggestion for an unknown name: the alias nearest by
/// optimal-string-alignment distance (OfficeCLI's `EditDistance.Damerau`)
/// within `max(2, len / 3)`, several when they tie (OfficeCLI offers none on
/// a tie), or the vocabulary's element names when nothing is near.
pub(crate) fn did_you_mean(name: &str, vocabulary: &Vocabulary) -> String {
    let wanted: Vec<char> = name
        .chars()
        .take(128)
        .map(|c| c.to_ascii_lowercase())
        .collect();

    let threshold = (wanted.len() / 3).max(2);
    let mut best = usize::MAX;
    let mut nearest: Vec<(&str, &str)> = Vec::new();
    for (alias, canonical) in vocabulary.spellings() {
        let candidate: Vec<char> = alias.chars().collect();
        let distance = distance(&wanted, &candidate);
        if distance > threshold || distance > best {
            continue;
        }
        if distance < best {
            best = distance;
            nearest.clear();
        }
        if !nearest.iter().any(|(_, seen)| *seen == canonical) {
            nearest.push((alias, canonical));
        }
    }
    match nearest.as_slice() {
        [] => {
            let mut names: Vec<&str> = vocabulary
                .elements
                .iter()
                .map(|rule| rule.name)
                .filter(|element| *element != vocabulary.root)
                .collect();
            names.sort_by_key(|element| element.to_ascii_lowercase());
            format!("known elements: {}", names.join(", "))
        },
        [(alias, _)] => format!("did you mean \"{alias}\"?"),
        several => {
            let names: Vec<String> = several
                .iter()
                .map(|(alias, _)| format!("\"{alias}\""))
                .collect();
            format!("did you mean one of {}?", names.join(", "))
        },
    }
}

/// Optimal string alignment distance: insertions, deletions, substitutions
/// and transpositions of adjacent characters each cost one.
pub(crate) fn distance(left: &[char], right: &[char]) -> usize {
    let width = right.len() + 1;
    let mut table = vec![0_usize; (left.len() + 1) * width];
    for (row, cell) in table.iter_mut().step_by(width).enumerate() {
        *cell = row;
    }
    for (column, cell) in table.iter_mut().take(width).enumerate() {
        *cell = column;
    }
    for row in 1..=left.len() {
        for column in 1..=right.len() {
            let cost = usize::from(left[row - 1] != right[column - 1]);
            let mut value = (table[(row - 1) * width + column] + 1)
                .min(table[row * width + column - 1] + 1)
                .min(table[(row - 1) * width + column - 1] + cost);
            if row > 1
                && column > 1
                && left[row - 1] == right[column - 2]
                && left[row - 2] == right[column - 1]
            {
                value = value.min(table[(row - 2) * width + column - 2] + 1);
            }
            table[row * width + column] = value;
        }
    }
    table[left.len() * width + right.len()]
}
