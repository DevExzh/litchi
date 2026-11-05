//! Element paths: `/body/p[3]/r[1]`, `/Sheet1/A1:C5`, `/slide[2]/notes`.

use std::borrow::Cow;
use std::fmt::{self, Display as _, Write as _};
use std::str::FromStr;

use super::axis::{Axis, CellRange, CellRef};
use super::cursor::{Cursor, Result};
use super::error::{ErrorKind, Expected, Grammar, Limit, ParseError};
use super::limits;
use super::predicate::{CompareOp, Condition, Expr, Key, Predicate, Predicates, Value};
use super::store::{Inline, Placeholder};
use super::text::{is_plain_name, is_quote, write_quoted};

/// Inline step capacity: the deepest common OfficeCLI forms
/// (`/body/tbl[1]/tr[2]/tc[3]/p[1]/r[1]`) fit without a heap allocation.
const INLINE_STEPS: usize = 6;

/// A parsed element path.
///
/// Absolute paths start at the document root (`/body/p[1]`); relative paths
/// (`p[@paraId=1A2B3C4D]`) are anchors resolved against a parent path. The
/// root path `/` has no steps.
///
/// [`Path::parse`] borrows segment names and values from the input; short
/// paths allocate nothing. Use [`Path::into_owned`] or [`str::parse`] for an
/// owned `Path<'static>`.
///
/// ```
/// use litchi_core::selector::{Axis, Path};
///
/// let path = Path::parse("/header[1]/p[2]")?;
/// assert_eq!(path.steps()[0].axis(), Some(Axis::Header));
/// assert_eq!(path.to_string(), "/header[1]/p[2]");
///
/// let owned: Path<'static> = "/Sheet1/A1:C5".parse()?;
/// assert!(owned.steps()[1].cell_range().is_some());
/// # Ok::<(), litchi_core::selector::ParseError>(())
/// ```
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Path<'a> {
    absolute: bool,
    steps: Inline<Step<'a>, INLINE_STEPS>,
}

impl<'a> Path<'a> {
    /// The root path `/`.
    #[must_use]
    pub fn root() -> Self {
        Self {
            absolute: true,
            steps: Inline::new(),
        }
    }

    /// Creates an absolute path from steps.
    ///
    /// Steps beyond [`limits::MAX_STEPS`] are accepted here; only parsing and
    /// [`Path::push`] enforce the bound.
    #[must_use]
    pub fn from_steps(steps: impl IntoIterator<Item = Step<'a>>) -> Self {
        Self {
            absolute: true,
            steps: steps.into_iter().collect(),
        }
    }

    /// Creates a relative path from steps (an anchor such as
    /// `p[@paraId=1A2B3C4D]`).
    #[must_use]
    pub fn relative(steps: impl IntoIterator<Item = Step<'a>>) -> Self {
        Self {
            absolute: false,
            steps: steps.into_iter().collect(),
        }
    }

    /// Parses an absolute path, borrowing names and values from `input`.
    ///
    /// # Errors
    ///
    /// Returns a [`ParseError`] with code `invalid_path` (or
    /// `max_depth_exceeded` for step and nesting limits) and the byte offset
    /// of the failure.
    #[inline]
    pub fn parse(input: &'a str) -> std::result::Result<Self, ParseError> {
        // Inlined so the fast path's result is built where the caller keeps
        // it rather than copied there (record 0326); the general parser
        // stays out of line.
        let mut path = Path::root();
        if parse_plain(input, &mut path) {
            return Ok(path);
        }
        parse_general(input)
    }

    /// Parses a relative anchor path such as `p[@paraId=1A2B3C4D]` or
    /// `tr[2]/tc[1]` (no leading `/`).
    ///
    /// # Errors
    ///
    /// As [`Path::parse`].
    pub fn parse_relative(input: &'a str) -> std::result::Result<Self, ParseError> {
        let mut cursor = Cursor::new(input, Grammar::Path)?;
        parse_path(&mut cursor, false)
    }

    /// Returns `true` for paths that start at the document root.
    #[must_use]
    pub const fn is_absolute(&self) -> bool {
        self.absolute
    }

    /// Returns `true` for the root path `/`.
    #[must_use]
    pub fn is_root(&self) -> bool {
        self.absolute && self.steps.len() == 0
    }

    /// Returns the steps in order.
    #[must_use]
    pub fn steps(&self) -> &[Step<'a>] {
        self.steps.as_slice()
    }

    /// Returns the final step, if any.
    #[must_use]
    pub fn last(&self) -> Option<&Step<'a>> {
        self.steps.as_slice().last()
    }

    /// Appends a step.
    ///
    /// # Errors
    ///
    /// Returns `max_depth_exceeded` when the path already has
    /// [`limits::MAX_STEPS`] steps.
    pub fn push(&mut self, step: Step<'a>) -> std::result::Result<(), ParseError> {
        if self.steps.len() >= limits::MAX_STEPS {
            return Err(ParseError::new(
                Grammar::Path,
                0,
                ErrorKind::LimitExceeded(Limit::Steps),
            ));
        }
        self.steps.push(step);
        Ok(())
    }

    /// Returns the first story or part axis named by the path, with its step
    /// index (`/header[1]/p[1]` -> `(0, Header)`;
    /// `/slide[2]/notes` -> `(0, Slide)`; use [`Path::axes`] for all).
    #[must_use]
    pub fn story(&self) -> Option<(usize, Axis)> {
        self.axes().next()
    }

    /// Iterates every step that names a story or part axis.
    pub fn axes(&self) -> impl Iterator<Item = (usize, Axis)> + '_ {
        self.steps
            .as_slice()
            .iter()
            .enumerate()
            .filter_map(|(index, step)| step.axis().map(|axis| (index, axis)))
    }

    /// Returns `true` when every predicate addresses by position or stable
    /// identity (see [`Predicate::is_structural`]); `false` when some step
    /// filters by content (`/Sheet1/cell[value>5000]`).
    #[must_use]
    pub fn is_structural(&self) -> bool {
        self.steps.as_slice().iter().all(|step| {
            step.predicates
                .as_slice()
                .iter()
                .all(Predicate::is_structural)
        })
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Path<'static> {
        Path {
            absolute: self.absolute,
            steps: self.steps.map(Step::into_owned),
        }
    }
}

impl Default for Path<'_> {
    fn default() -> Self {
        Self::root()
    }
}

impl FromStr for Path<'static> {
    type Err = ParseError;

    fn from_str(input: &str) -> std::result::Result<Self, Self::Err> {
        Path::parse(input).map(Path::into_owned)
    }
}

impl fmt::Display for Path<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.is_root() {
            return f.write_char('/');
        }
        for (index, step) in self.steps.as_slice().iter().enumerate() {
            if self.absolute || index > 0 {
                f.write_char('/')?;
            }
            step.fmt(f)?;
        }
        Ok(())
    }
}

/// One path segment: a name plus zero or more `[...]` predicates.
///
/// Names are kept verbatim (case and spelling preserved); comparing them to
/// document element names, aliases (`paragraph` for `p`), sheet names, and
/// style IDs is resolution's job. A segment that needs `/`, `[`, `]`, or a
/// leading quote is written quoted: `/"a/b"/A1`.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Step<'a> {
    name: Cow<'a, str>,
    predicates: Predicates<'a>,
}

impl<'a> Step<'a> {
    /// Creates a step with no predicates.
    pub fn new(name: impl Into<Cow<'a, str>>) -> Self {
        Self {
            name: name.into(),
            predicates: Predicates::new(),
        }
    }

    /// Adds a predicate (builder style).
    #[must_use]
    pub fn with(mut self, predicate: Predicate<'a>) -> Self {
        self.predicates.push(predicate);
        self
    }

    /// Returns the segment name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the predicates in order.
    #[must_use]
    pub fn predicates(&self) -> &[Predicate<'a>] {
        self.predicates.as_slice()
    }

    /// Returns the step's only predicate when it is a positional index.
    #[must_use]
    pub fn index(&self) -> Option<super::Position> {
        match self.predicates() {
            [Predicate::Index(position)] => Some(*position),
            _ => None,
        }
    }

    /// Classifies the name as a story or part axis (case-insensitive).
    #[must_use]
    pub fn axis(&self) -> Option<Axis> {
        Axis::classify(&self.name)
    }

    /// Reads a bare `A1` cell-reference segment.
    #[must_use]
    pub fn cell_ref(&self) -> Option<CellRef> {
        if self.predicates.is_empty() {
            CellRef::parse(&self.name)
        } else {
            None
        }
    }

    /// Reads a bare `A1` or `A1:C5` cell-range segment (a single cell is a
    /// one-cell range).
    #[must_use]
    pub fn cell_range(&self) -> Option<CellRange> {
        if self.predicates.is_empty() {
            CellRange::parse(&self.name)
        } else {
            None
        }
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Step<'static> {
        Step {
            name: Cow::Owned(self.name.into_owned()),
            predicates: self.predicates.map(Predicate::into_owned),
        }
    }

    pub(super) fn fmt_with(&self, f: &mut fmt::Formatter<'_>, in_query_scope: bool) -> fmt::Result {
        let plain = is_plain_name(&self.name)
            && !(in_query_scope && self.name.bytes().any(|b| b == b'(' || b == b')'));
        if plain {
            f.write_str(&self.name)?;
        } else {
            write_quoted(f, &self.name)?;
        }
        for predicate in self.predicates() {
            predicate.fmt(f)?;
        }
        Ok(())
    }
}

impl Placeholder for Step<'_> {
    fn placeholder() -> Self {
        Self::new("")
    }
}

impl fmt::Display for Step<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.fmt_with(f, false)
    }
}

/// Parses the common absolute forms in one pass over the bytes (records
/// 0302 and 0326): `/` followed by plain segment names, each with zero or
/// more predicates of the forms `[N]`, `[last()]`, `[key]`, `[@key]`,
/// `[key=value]` and `[@key=value]` written without whitespace, quotes or
/// escapes (`/body/p[3]`, `/sheet[1]/table[1]/B3`,
/// `/body/p[@paraId=1A2B3C4D]`, `/slide[1]/shape[@id=4]/paragraph[2]`).
///
/// Fills `path` (the root path on entry) and returns `true`; returns
/// `false`, leaving `path` unspecified, for anything else, valid or not: a
/// quoted name, any other predicate, whitespace inside brackets, a
/// non-ASCII key, an index of zero or one that overflows, a step,
/// predicate or byte limit, or a malformed path. The general parser then
/// builds the path or reports the error, so every error, offset and
/// message is the general parser's. For every input this accepts, the
/// general parser builds an equal path (the tests
/// `plain_fast_path_matches_the_general_parser` and, over every path of the
/// test suites and goldens, `litchi-ops`' `path_parser_differential`). The
/// path is filled in place: a parsed path is several hundred bytes, and
/// returning it through an `Option` copied it once more.
fn parse_plain<'a>(input: &'a str, path: &mut Path<'a>) -> bool {
    parse_plain_steps(input, path).is_some()
}

/// Byte classes of the fast path (record 0326): one table lookup per byte
/// instead of a chain of comparisons.
const DELIMITER: u8 = 1;
const CONTROL: u8 = 2;
const KEY: u8 = 4;
/// Ends or invalidates a bare predicate value on the fast path: `]`, `)`,
/// `[`, whitespace, quotes, `*` and the operator bytes.
const VALUE_STOP: u8 = 8;

const fn byte_classes() -> [u8; 256] {
    let mut table = [0_u8; 256];
    let mut byte = 0;
    while byte < 256 {
        #[allow(clippy::cast_possible_truncation, reason = "byte < 256")]
        let b = byte as u8;
        let mut class = 0;
        if matches!(b, b'/' | b'[' | b']') {
            class |= DELIMITER;
        }
        if b.is_ascii_control() {
            class |= CONTROL | VALUE_STOP;
        }
        if b.is_ascii_alphanumeric() || b == b'.' || b == b'_' {
            class |= KEY;
        }
        if matches!(
            b,
            b']' | b')'
                | b'['
                | b'"'
                | b'\''
                | b'*'
                | b'='
                | b'!'
                | b'~'
                | b'<'
                | b'>'
                | b'\\'
                | b' '
                | b'\t'
                | b'\n'
                | b'\r'
                | 0x0c
        ) {
            class |= VALUE_STOP;
        }
        table[byte] = class;
        byte += 1;
    }
    table
}

static CLASSES: [u8; 256] = byte_classes();

#[inline(always)]
fn class(byte: u8) -> u8 {
    CLASSES[usize::from(byte)]
}

fn parse_plain_steps<'a>(input: &'a str, path: &mut Path<'a>) -> Option<()> {
    let bytes = input.as_bytes();
    if bytes.first() != Some(&b'/') || bytes.len() > limits::MAX_INPUT_BYTES {
        return None;
    }
    let mut at = 1;
    if at == bytes.len() {
        return Some(());
    }
    loop {
        let start = at;
        if is_quote(bytes[at]) {
            return None;
        }
        while at < bytes.len() && class(bytes[at]) & (DELIMITER | CONTROL) == 0 {
            at += 1;
        }
        // A control byte is the general parser's error.
        if at == start
            || path.steps.len() >= limits::MAX_STEPS
            || bytes
                .get(at)
                .is_some_and(|&byte| class(byte) & CONTROL != 0)
        {
            return None;
        }
        // `at` is at an ASCII delimiter or the end: a char boundary.
        let mut step = Step::new(&input[start..at]);
        while bytes.get(at) == Some(&b'[') {
            if step.predicates.len() >= limits::MAX_PREDICATES {
                return None;
            }
            at += 1;
            if bytes.get(at).is_some_and(u8::is_ascii_digit) {
                // `[N]`, the common case, built in place.
                let mut value: usize = 0;
                while let Some(&byte) = bytes.get(at).filter(|byte| byte.is_ascii_digit()) {
                    value = value
                        .checked_mul(10)?
                        .checked_add(usize::from(byte - b'0'))?;
                    at += 1;
                }
                if bytes.get(at) != Some(&b']') {
                    return None;
                }
                at += 1;
                step.predicates
                    .push(Predicate::Index(super::Position::from_one_based(value)?));
            } else {
                let predicate = plain_predicate(input, &mut at)?;
                step.predicates.push(predicate);
            }
        }
        path.steps.push(step);
        match bytes.get(at) {
            None => return Some(()),
            Some(b'/') if at + 1 < bytes.len() => at += 1,
            Some(_) => return None,
        }
    }
}

/// One fast-path predicate other than `[N]`; `at` is just past `[` and
/// ends just past `]`.
fn plain_predicate<'a>(input: &'a str, at: &mut usize) -> Option<Predicate<'a>> {
    let bytes = input.as_bytes();
    let first = *bytes.get(*at)?;
    if first == b'l' || first == b'L' {
        if let Some(rest) = bytes.get(*at..*at + 7)
            && rest[..6].eq_ignore_ascii_case(b"last()")
            && rest[6] == b']'
        {
            *at += 7;
            return Some(Predicate::Last);
        }
    }
    let attribute = first == b'@';
    if attribute {
        *at += 1;
    }
    let key_start = *at;
    while bytes.get(*at).is_some_and(|&byte| class(byte) & KEY != 0) {
        *at += 1;
    }
    // An empty key, a key that starts like an index (`[1a]`), or a key
    // followed by anything but `]` or `=` takes the general parser.
    if *at == key_start || bytes[key_start].is_ascii_digit() {
        return None;
    }
    let key = Key::from_parts(Cow::Borrowed(&input[key_start..*at]), attribute);
    match bytes.get(*at)? {
        b']' => {
            *at += 1;
            Some(Predicate::Filter(Expr::Condition(Condition::Exists(key))))
        },
        b'=' => {
            *at += 1;
            let value_start = *at;
            while bytes
                .get(*at)
                .is_some_and(|&byte| class(byte) & VALUE_STOP == 0)
            {
                *at += 1;
            }
            if *at == value_start || bytes.get(*at) != Some(&b']') {
                return None;
            }
            // `]` is ASCII: a char boundary.
            let value = Value::Text(Cow::Borrowed(&input[value_start..*at]));
            *at += 1;
            Some(Predicate::Filter(Expr::Condition(Condition::Compare {
                key,
                op: CompareOp::Eq,
                value,
            })))
        },
        _ => None,
    }
}

/// [`Path::parse`] for every input the fast path does not take.
#[inline(never)]
fn parse_general(input: &str) -> std::result::Result<Path<'_>, ParseError> {
    let mut cursor = Cursor::new(input, Grammar::Path)?;
    parse_path(&mut cursor, true)
}

/// Parses a path over the cursor window, which must be consumed entirely.
pub(super) fn parse_path<'a>(cursor: &mut Cursor<'a>, absolute: bool) -> Result<Path<'a>> {
    if cursor.at_end() {
        return Err(cursor.error(ErrorKind::Empty));
    }
    let mut path = Path {
        absolute,
        steps: Inline::new(),
    };
    if absolute {
        cursor.expect(b'/', Expected::Root)?;
        if cursor.at_end() {
            return Ok(path);
        }
    }
    loop {
        if cursor.peek() == Some(b'/') {
            return Err(cursor.error(ErrorKind::EmptySegment));
        }
        if path.steps.len() >= limits::MAX_STEPS {
            return Err(cursor.error(ErrorKind::LimitExceeded(Limit::Steps)));
        }
        path.steps.push(step(cursor)?);
        match cursor.peek() {
            None => return Ok(path),
            Some(b'/') => {
                cursor.bump();
                if cursor.at_end() {
                    return Err(cursor.error_at(cursor.pos() - 1, ErrorKind::EmptySegment));
                }
            },
            Some(_) => return Err(cursor.unexpected(Expected::StepEnd)),
        }
    }
}

fn step<'a>(cursor: &mut Cursor<'a>) -> Result<Step<'a>> {
    let name = match cursor.peek() {
        Some(b) if is_quote(b) => cursor.quoted()?,
        _ => {
            let start = cursor.pos();
            while let Some(b) = cursor.peek() {
                match b {
                    b'/' | b'[' | b']' => break,
                    _ if b.is_ascii_control() => {
                        return Err(cursor.error(ErrorKind::ControlCharacter));
                    },
                    _ => cursor.bump(),
                }
            }
            if cursor.pos() == start {
                return Err(cursor.unexpected(Expected::SegmentName));
            }
            Cow::Borrowed(&cursor.src()[start..cursor.pos()])
        },
    };
    let mut predicates = Predicates::new();
    cursor.predicates(&mut predicates, 0)?;
    Ok(Step { name, predicates })
}

impl fmt::Debug for Step<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Step { .. }")
    }
}

#[cfg(test)]
mod plain_tests {
    use super::*;

    fn general(input: &str) -> std::result::Result<Path<'_>, ParseError> {
        let mut cursor = Cursor::new(input, Grammar::Path)?;
        parse_path(&mut cursor, true)
    }

    fn plain(input: &str) -> Option<Path<'_>> {
        let mut path = Path::root();
        parse_plain(input, &mut path).then_some(path)
    }

    fn check(input: &str) -> bool {
        let Some(fast) = plain(input) else {
            return false;
        };
        let slow = general(input).unwrap_or_else(|error| {
            panic!("{input:?}: the fast path accepted what the parser refuses: {error}")
        });
        assert_eq!(fast, slow, "{input:?}");
        assert_eq!(fast.to_string(), slow.to_string(), "{input:?}");
        for (fast, slow) in fast.steps().iter().zip(slow.steps()) {
            assert_eq!(fast.predicates(), slow.predicates(), "{input:?}");
        }
        true
    }

    /// The quoted strings of a text file that start with `/`, with the
    /// common `\"` and `\\` escapes reversed.
    fn quoted_paths(text: &str, out: &mut std::collections::BTreeSet<String>) {
        let bytes = text.as_bytes();
        let mut at = 0;
        while let Some(found) = bytes[at..].iter().position(|&byte| byte == b'"') {
            let start = at + found + 1;
            let mut end = start;
            let mut value = String::new();
            let mut closed = false;
            while let Some(&byte) = bytes.get(end) {
                match byte {
                    b'"' => {
                        closed = true;
                        break;
                    },
                    b'\n' => break,
                    b'\\' if matches!(bytes.get(end + 1), Some(b'"' | b'\\')) => {
                        value.push(char::from(bytes[end + 1]));
                        end += 2;
                        continue;
                    },
                    _ => {},
                }
                // Copy whole characters.
                let width = text[end..].chars().next().map_or(1, char::len_utf8);
                value.push_str(&text[end..end + width]);
                end += width;
            }
            if closed && value.starts_with('/') && value.len() <= 512 {
                out.insert(value);
            }
            at = if closed { end + 1 } else { end.max(start) };
            if at >= bytes.len() {
                break;
            }
        }
    }

    fn collect(dir: &std::path::Path, out: &mut std::collections::BTreeSet<String>) {
        let Ok(entries) = std::fs::read_dir(dir) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            let Ok(kind) = entry.file_type() else {
                continue;
            };
            if kind.is_dir() {
                if path.file_name().is_some_and(|name| name != "target") {
                    collect(&path, out);
                }
            } else if kind.is_file()
                && path.extension().is_some_and(|extension| {
                    ["rs", "json", "jsonl", "tsv", "md", "txt", "sh", "py"]
                        .iter()
                        .any(|known| extension == *known)
                })
                && let Ok(text) = std::fs::read_to_string(&path)
            {
                quoted_paths(&text, out);
            }
        }
    }

    /// Record 0326: the fast path against the general parser over every
    /// quoted `/…` string of the workspace's sources, tests and goldens (the
    /// OfficeCLI corpus, the query corpus, the capability, registry and
    /// node-view goldens, the CRUD checklist), each also cut at every
    /// character boundary: equal paths, and byte-identical errors (kind,
    /// offset and message) from `Path::parse` and the general parser.
    #[test]
    fn path_parser_differential() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let mut inputs = std::collections::BTreeSet::new();
        collect(&root.join("crates"), &mut inputs);
        collect(&root.join("docs/unified-ops/results/0289"), &mut inputs);
        assert!(
            inputs.len() > 2_000,
            "only {} path strings found",
            inputs.len()
        );
        let (mut accepted, mut compared) = (0_usize, 0_usize);
        for input in &inputs {
            for (cut, _) in input.char_indices().skip(1).chain([(input.len(), ' ')]) {
                let text = &input[..cut];
                compared += 1;
                if check(text) {
                    accepted += 1;
                }
                match (Path::parse(text), general(text)) {
                    (Ok(fast), Ok(slow)) => {
                        assert_eq!(fast, slow, "{text:?}");
                        assert_eq!(fast.to_string(), slow.to_string(), "{text:?}");
                    },
                    (Err(fast), Err(slow)) => {
                        assert_eq!(fast, slow, "{text:?}");
                        assert_eq!(fast.to_string(), slow.to_string(), "{text:?}");
                    },
                    (fast, slow) => panic!("{text:?}: {fast:?} against {slow:?}"),
                }
            }
        }
        eprintln!(
            "path_parser_differential: {} strings, {compared} inputs, {accepted} on the fast path",
            inputs.len()
        );
        assert!(accepted > 1_000, "{accepted} of {compared}");
    }

    #[test]
    fn plain_fast_path_matches_the_general_parser() {
        // Every string of up to five bytes over the bytes the grammars
        // treat specially, plus name and digit bytes.
        let alphabet = b"/aB[]10 \"'\t*@=";
        let mut accepted = 0;
        let mut input = Vec::new();
        // Every string of length 1..=5, depth first.
        fn visit(alphabet: &[u8], input: &mut Vec<u8>, depth: usize, accepted: &mut usize) {
            if depth == 0 {
                return;
            }
            for &byte in alphabet {
                input.push(byte);
                let text = std::str::from_utf8(input).expect("ascii");
                if check(text) {
                    *accepted += 1;
                }
                visit(alphabet, input, depth - 1, accepted);
                input.pop();
            }
        }
        visit(alphabet, &mut input, 5, &mut accepted);
        assert!(accepted > 100, "{accepted}");
        for input in [
            "/",
            "/body/p[1]",
            "/body/p[3]/r[12]",
            "/body/tbl[1]/tr[2]/tc[3]/p[1]/r[1]",
            "/body/tbl[1]/tr[2]/tc[3]/p[1]/r[1]/x[2]/y",
            "/sheet[1]/table[1]/B3",
            "/Sheet1/A1:C5",
            "/slide[2]/notes",
            "/Größe/ü[007]",
            "/a[18446744073709551615]",
            // Record 0326: attribute, bare-key and `last()` predicates.
            "/a[last()]",
            "/a[LaSt()]",
            "/a[@id=1]",
            "/body/p[@paraId=1A2B3C4D]",
            "/slide[1]/shape[@id=4]/paragraph[2]",
            "/slide[1]/shape[@name=Title.1_x]",
            "/Sheet1/col[B]",
            "/slide[1]/placeholder[ctrTitle]",
            "/a[not]",
            "/a[or]",
            "/a[@k]",
            "/a[k=v]",
            "/a[k=ü]",
            "/a[k=a(b]",
            "/a[k=1,2]",
            "/a[k=a:b]",
            "/a[1][2]",
            "/a[@id=1][2][last()]",
            "/a[l]",
            "/a[last]",
        ] {
            assert!(check(input), "{input:?}");
        }
        for refused in [
            "",
            "body",
            "//",
            "/a/",
            "/a[0]",
            "/a[18446744073709551616]",
            "/a[ 1]",
            "/\"a/b\"/A1",
            "/a[last( )]",
            "/a[last()x]",
            "/a[@]",
            "/a[@=1]",
            "/a[1a]",
            "/a[-1]",
            "/a[ü]",
            "/a[k==v]",
            "/a[k=~v]",
            "/a[k!=v]",
            "/a[k~=v]",
            "/a[k>v]",
            "/a[k=v w]",
            "/a[k=\"v\"]",
            "/a[k=r'v']",
            "/a[k=v*]",
            "/a[k=v)]",
            "/a[k=]",
            "/a[k=v",
            "/a[k=v]]",
            "/a[k and j]",
            "/a[not(k)]",
            "/a[k=\u{7}]",
            "/a]",
            "/a\u{7}",
        ] {
            assert!(plain(refused).is_none(), "{refused:?}");
            assert_eq!(
                Path::parse(refused).map(|path| path.to_string()),
                general(refused).map(|path| path.to_string()),
                "{refused:?}"
            );
        }
        let deep = "/a".repeat(limits::MAX_STEPS + 1);
        assert!(plain(&deep).is_none());
        assert!(Path::parse(&deep).is_err());
        let limit = "/a".repeat(limits::MAX_STEPS);
        assert!(check(&limit));
        let many = format!("/a{}", "[1]".repeat(limits::MAX_PREDICATES + 1));
        assert!(plain(&many).is_none());
        assert!(Path::parse(&many).is_err());
        let most = format!("/a{}", "[1]".repeat(limits::MAX_PREDICATES));
        assert!(check(&most));
        let long = format!("/{}", "a".repeat(limits::MAX_INPUT_BYTES));
        assert!(plain(&long).is_none());
    }
}
