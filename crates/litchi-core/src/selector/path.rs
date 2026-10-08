//! Element paths: `/body/p[3]/r[1]`, `/Sheet1/A1:C5`, `/slide[2]/notes`.

use std::borrow::Cow;
use std::fmt::{self, Display as _, Write as _};
use std::str::FromStr;

use super::axis::{Axis, CellRange, CellRef};
use super::cursor::{Cursor, Result};
use super::error::{ErrorKind, Expected, Grammar, Limit, ParseError};
use super::limits;
use super::predicate::{Predicate, Predicates};
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
    pub fn parse(input: &'a str) -> std::result::Result<Self, ParseError> {
        let mut cursor = Cursor::new(input, Grammar::Path)?;
        parse_path(&mut cursor, true)
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
