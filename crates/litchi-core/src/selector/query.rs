//! CSS-like query selectors: `paragraph[style=Heading1] > run[font!=Arial]`,
//! `Sheet1!row[Salary>5000 and Region=EMEA]`, `shape:contains("Revenue")`.

use std::borrow::Cow;
use std::fmt::{self, Display as _, Write as _};
use std::str::FromStr;

use super::axis::CellRange;
use super::cursor::{Cursor, Result};
use super::error::{ErrorKind, Expected, Grammar, Limit, ParseError};
use super::limits;
use super::path::{Path, parse_path};
use super::predicate::{Predicate, Predicates};
use super::store::OneOrMany;
use super::text::{
    is_ident_byte, is_plain_contains, is_plain_sheet, is_quote, is_ws, starts_with_ci,
    write_quoted, write_sheet_quoted,
};

/// A parsed query: an optional scope and one or more comma-separated
/// alternatives (a union).
///
/// The AST is immutable, `Send + Sync`, and reusable: parse once, then
/// evaluate against any number of nodes. Element names and keys are kept
/// verbatim; case-insensitive matching and aliasing are resolution's job.
///
/// ```
/// use litchi_core::selector::{Combinator, Query, Scope};
///
/// let query = Query::parse("paragraph[style=Normal] > run[font!=Arial]")?;
/// let complex = &query.alternatives()[0];
/// assert_eq!(complex.compounds().count(), 2);
/// assert_eq!(complex.rest()[0].0, Combinator::Child);
///
/// let scoped = Query::parse("'My Data (2024)'!row[Amount>5000]")?;
/// assert_eq!(scoped.scope(), Some(&Scope::Sheet("My Data (2024)".into())));
/// assert_eq!(scoped.to_string(), "'My Data (2024)'!row[Amount>5000]");
/// # Ok::<(), litchi_core::selector::ParseError>(())
/// ```
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Query<'a> {
    scope: Option<Scope<'a>>,
    alternatives: OneOrMany<Complex<'a>>,
}

/// Where a query applies.
#[derive(Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Scope<'a> {
    /// A leading absolute path: `/Sheet1/cell[...]` has scope `/Sheet1`;
    /// `/body` has the root scope. Always rendered absolute. Boxed because a
    /// path keeps its first steps inline.
    Path(Box<Path<'a>>),
    /// An Excel-style sheet prefix: `Sheet1!B[value!=0]`,
    /// `'Q1 Data'!row[...]`.
    Sheet(Cow<'a, str>),
}

impl<'a> Query<'a> {
    /// Creates an unscoped query from alternatives.
    #[must_use]
    pub fn new(alternatives: impl IntoIterator<Item = Complex<'a>>) -> Self {
        Self {
            scope: None,
            alternatives: alternatives.into_iter().collect(),
        }
    }

    /// Sets the scope (builder style).
    #[must_use]
    pub fn with_scope(mut self, scope: Scope<'a>) -> Self {
        self.scope = Some(scope);
        self
    }

    /// Parses a query, borrowing names and values from `input`. Leading and
    /// trailing ASCII whitespace is ignored.
    ///
    /// # Errors
    ///
    /// Returns a [`ParseError`] with code `invalid_selector` (or
    /// `max_depth_exceeded` for nesting limits) and the byte offset of the
    /// failure. A scope path that fails reports code `invalid_selector` too.
    pub fn parse(input: &'a str) -> std::result::Result<Self, ParseError> {
        let mut cursor = Cursor::new(input, Grammar::Query)?;
        cursor.skip_ws();
        let trimmed_end = input
            .trim_end_matches(|c: char| c.is_ascii() && is_ws(c as u8))
            .len();
        cursor.set_end(trimmed_end.max(cursor.pos()));
        if cursor.at_end() {
            return Err(cursor.error(ErrorKind::Empty));
        }
        let scope = scope(&mut cursor)?;
        if cursor.at_end() {
            return Err(cursor.unexpected(Expected::Compound));
        }
        let alternatives = selector_list(&mut cursor, false)?;
        if !cursor.at_end() {
            return Err(cursor.unexpected(Expected::CombinatorOrEnd));
        }
        // Box the scope path only once the whole query is known to be valid,
        // so rejected input never allocates.
        let scope = scope.map(|scope| match scope {
            ParsedScope::Path(path) => Scope::Path(Box::new(path)),
            ParsedScope::Sheet(name) => Scope::Sheet(name),
        });
        Ok(Self {
            scope,
            alternatives,
        })
    }

    /// Returns the scope, if any.
    #[must_use]
    pub fn scope(&self) -> Option<&Scope<'a>> {
        self.scope.as_ref()
    }

    /// Returns the comma-separated alternatives (at least one when parsed).
    #[must_use]
    pub fn alternatives(&self) -> &[Complex<'a>] {
        self.alternatives.as_slice()
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Query<'static> {
        Query {
            scope: self.scope.map(Scope::into_owned),
            alternatives: self.alternatives.map(Complex::into_owned),
        }
    }
}

impl Scope<'_> {
    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Scope<'static> {
        match self {
            Self::Path(path) => Scope::Path(Box::new(path.into_owned())),
            Self::Sheet(name) => Scope::Sheet(Cow::Owned(name.into_owned())),
        }
    }
}

impl FromStr for Query<'static> {
    type Err = ParseError;

    fn from_str(input: &str) -> std::result::Result<Self, Self::Err> {
        Query::parse(input).map(Query::into_owned)
    }
}

impl fmt::Display for Query<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.scope {
            None => {},
            Some(Scope::Path(path)) => {
                for step in path.steps() {
                    f.write_char('/')?;
                    step.fmt_with(f, true)?;
                }
                f.write_char('/')?;
            },
            Some(Scope::Sheet(name)) => {
                if is_plain_sheet(name) {
                    f.write_str(name)?;
                } else {
                    write_sheet_quoted(f, name)?;
                }
                f.write_char('!')?;
            },
        }
        write_list(f, self.alternatives.as_slice())
    }
}

fn write_list(f: &mut fmt::Formatter<'_>, list: &[Complex<'_>]) -> fmt::Result {
    for (index, complex) in list.iter().enumerate() {
        if index > 0 {
            f.write_str(", ")?;
        }
        complex.fmt(f)?;
    }
    Ok(())
}

/// A chain of compounds joined by combinators: `slide > shape`,
/// `table tc`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Complex<'a> {
    first: Compound<'a>,
    rest: Vec<(Combinator, Compound<'a>)>,
}

impl<'a> Complex<'a> {
    /// Creates a single-compound selector.
    #[must_use]
    pub fn new(first: Compound<'a>) -> Self {
        Self {
            first,
            rest: Vec::new(),
        }
    }

    /// Appends a combinator and compound (builder style).
    #[must_use]
    pub fn then(mut self, combinator: Combinator, compound: Compound<'a>) -> Self {
        self.rest.push((combinator, compound));
        self
    }

    /// Returns the leftmost compound.
    #[must_use]
    pub fn first(&self) -> &Compound<'a> {
        &self.first
    }

    /// Returns the following `(combinator, compound)` links.
    #[must_use]
    pub fn rest(&self) -> &[(Combinator, Compound<'a>)] {
        &self.rest
    }

    /// Returns the rightmost compound: the elements the selector returns.
    #[must_use]
    pub fn subject(&self) -> &Compound<'a> {
        self.rest
            .last()
            .map_or(&self.first, |(_, compound)| compound)
    }

    /// Iterates all compounds left to right.
    pub fn compounds(&self) -> impl Iterator<Item = &Compound<'a>> + '_ {
        std::iter::once(&self.first).chain(self.rest.iter().map(|(_, compound)| compound))
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Complex<'static> {
        Complex {
            first: self.first.into_owned(),
            rest: self
                .rest
                .into_iter()
                .map(|(combinator, compound)| (combinator, compound.into_owned()))
                .collect(),
        }
    }
}

impl fmt::Display for Complex<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.first.fmt(f)?;
        for (combinator, compound) in &self.rest {
            f.write_str(match combinator {
                Combinator::Child => " > ",
                Combinator::Descendant => " ",
            })?;
            compound.fmt(f)?;
        }
        Ok(())
    }
}

/// How two compounds relate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Combinator {
    /// `a > b`: `b` is a direct child of `a`.
    Child,
    /// `a b`: `b` is a descendant of `a`.
    Descendant,
}

/// One element test: optional type, bracket predicates, and pseudo-classes,
/// all of which must hold.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Compound<'a> {
    element: Element<'a>,
    predicates: Predicates<'a>,
    pseudos: OneOrMany<Pseudo<'a>>,
}

impl<'a> Compound<'a> {
    /// Creates a compound with an element test and no predicates.
    #[must_use]
    pub fn new(element: Element<'a>) -> Self {
        Self {
            element,
            predicates: Predicates::new(),
            pseudos: OneOrMany::new(),
        }
    }

    /// Adds a bracket predicate (builder style).
    #[must_use]
    pub fn with(mut self, predicate: Predicate<'a>) -> Self {
        self.predicates.push(predicate);
        self
    }

    /// Adds a pseudo-class (builder style).
    #[must_use]
    pub fn with_pseudo(mut self, pseudo: Pseudo<'a>) -> Self {
        self.pseudos.push(pseudo);
        self
    }

    /// Returns the element test.
    #[must_use]
    pub fn element(&self) -> &Element<'a> {
        &self.element
    }

    /// Returns the bracket predicates in order.
    #[must_use]
    pub fn predicates(&self) -> &[Predicate<'a>] {
        self.predicates.as_slice()
    }

    /// Returns the pseudo-classes in order.
    #[must_use]
    pub fn pseudos(&self) -> &[Pseudo<'a>] {
        self.pseudos.as_slice()
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Compound<'static> {
        Compound {
            element: self.element.into_owned(),
            predicates: self.predicates.map(Predicate::into_owned),
            pseudos: self.pseudos.map(Pseudo::into_owned),
        }
    }
}

impl fmt::Display for Compound<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.element {
            Element::Implicit => {},
            Element::Any => f.write_char('*')?,
            Element::Named(name) => f.write_str(name)?,
        }
        for predicate in self.predicates.as_slice() {
            predicate.fmt(f)?;
        }
        for pseudo in self.pseudos.as_slice() {
            pseudo.fmt(f)?;
        }
        Ok(())
    }
}

/// The element-type part of a compound.
#[derive(Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Element<'a> {
    /// No type written (`:has(formula)`, `[style=Title]`): any element.
    Implicit,
    /// `*`: every top-level element.
    Any,
    /// A type name (`paragraph`, `cell`, `B`) or cell reference / range
    /// (`A1`, `A1:C5`), verbatim. Names are ASCII letters, digits, `_`, `-`.
    Named(Cow<'a, str>),
}

impl Element<'_> {
    /// Returns the type name, if one was written.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        match self {
            Self::Named(name) => Some(name),
            Self::Implicit | Self::Any => None,
        }
    }

    /// Reads the name as an `A1` / `A1:C5` cell range.
    #[must_use]
    pub fn cell_range(&self) -> Option<CellRange> {
        self.name().and_then(CellRange::parse)
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Element<'static> {
        match self {
            Self::Implicit => Element::Implicit,
            Self::Any => Element::Any,
            Self::Named(name) => Element::Named(Cow::Owned(name.into_owned())),
        }
    }
}

/// A pseudo-class.
#[derive(Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Pseudo<'a> {
    /// `:contains(text)` (and OfficeCLI's shorthand `cell:text`, which parses
    /// to the same value): the element's text contains `text`.
    Contains(Cow<'a, str>),
    /// `:empty`: the element has no text.
    Empty,
    /// `:no-alt`: a picture without alternative text.
    NoAlt,
    /// `:has(selector)`: some descendant matches (`cell:has(formula)`).
    Has(Vec<Complex<'a>>),
    /// `:not(selector)`: the element does not match (`:not(:empty)`).
    Not(Vec<Complex<'a>>),
}

impl Pseudo<'_> {
    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Pseudo<'static> {
        match self {
            Self::Contains(text) => Pseudo::Contains(Cow::Owned(text.into_owned())),
            Self::Empty => Pseudo::Empty,
            Self::NoAlt => Pseudo::NoAlt,
            Self::Has(list) => Pseudo::Has(list.into_iter().map(Complex::into_owned).collect()),
            Self::Not(list) => Pseudo::Not(list.into_iter().map(Complex::into_owned).collect()),
        }
    }
}

impl fmt::Display for Pseudo<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Contains(text) => {
                f.write_str(":contains(")?;
                if is_plain_contains(text) {
                    f.write_str(text)?;
                } else {
                    write_quoted(f, text)?;
                }
                f.write_char(')')
            },
            Self::Empty => f.write_str(":empty"),
            Self::NoAlt => f.write_str(":no-alt"),
            Self::Has(list) => {
                f.write_str(":has(")?;
                write_list(f, list)?;
                f.write_char(')')
            },
            Self::Not(list) => {
                f.write_str(":not(")?;
                write_list(f, list)?;
                f.write_char(')')
            },
        }
    }
}

/// Finds the last `/` outside brackets, parentheses, and quoted spans. The
/// scan mirrors the parser's quoting rules: inside `[...]` and `(...)` every
/// quote opens a span; at top level only a quote that begins a segment does.
fn last_top_level_slash(bytes: &[u8], start: usize, end: usize) -> usize {
    let mut last = start;
    let mut brackets = 0usize;
    let mut parens = 0usize;
    let mut i = start;
    while i < end {
        let b = bytes[i];
        let nested = brackets > 0 || parens > 0;
        if is_quote(b) && (nested || (i > start && bytes[i - 1] == b'/')) {
            i += 1;
            while i < end && bytes[i] != b {
                i += if bytes[i] == b'\\' { 2 } else { 1 };
            }
            i += 1;
            continue;
        }
        match b {
            b'[' => brackets += 1,
            b']' => brackets = brackets.saturating_sub(1),
            b'(' if brackets == 0 => parens += 1,
            b')' if brackets == 0 => parens = parens.saturating_sub(1),
            b'/' if !nested => last = i,
            _ => {},
        }
        i += 1;
    }
    last
}

#[allow(
    clippy::large_enum_variant,
    reason = "a transient parse result, boxed into `Scope` only after the query validates"
)]
enum ParsedScope<'a> {
    Path(Path<'a>),
    Sheet(Cow<'a, str>),
}

fn scope<'a>(cursor: &mut Cursor<'a>) -> Result<Option<ParsedScope<'a>>> {
    let start = cursor.pos();
    match cursor.peek() {
        Some(b'/') => {
            let slash = last_top_level_slash(cursor.src().as_bytes(), start, cursor.end());
            let path_end = if slash == start { start + 1 } else { slash };
            let mut path_cursor = Cursor::window(cursor.src(), start, path_end, Grammar::Query);
            let path = parse_path(&mut path_cursor, true)?;
            cursor.set_pos(slash + 1);
            Ok(Some(ParsedScope::Path(path)))
        },
        Some(b'\'') => {
            // Excel quoting: '' is an escaped apostrophe.
            cursor.bump();
            let inner_start = cursor.pos();
            let mut escaped = false;
            loop {
                match cursor.peek() {
                    None => return Err(cursor.error_at(start, ErrorKind::UnterminatedQuote)),
                    Some(b'\'') if cursor.peek_at(1) == Some(b'\'') => {
                        escaped = true;
                        cursor.bump();
                        cursor.bump();
                    },
                    Some(b'\'') => break,
                    Some(_) => cursor.bump(),
                }
            }
            let inner = &cursor.src()[inner_start..cursor.pos()];
            cursor.bump();
            cursor.expect(b'!', Expected::SheetBang)?;
            let name = if escaped {
                Cow::Owned(inner.replace("''", "'"))
            } else {
                Cow::Borrowed(inner)
            };
            Ok(Some(ParsedScope::Sheet(name)))
        },
        _ => {
            let rest = cursor.rest();
            let stop = rest
                .iter()
                .position(|&b| {
                    matches!(b, b'!' | b'[' | b'(' | b':' | b',' | b'>' | b'/') || is_quote(b)
                })
                .unwrap_or(rest.len());
            if stop > 0 && rest.get(stop) == Some(&b'!') && rest.get(stop + 1) != Some(&b'=') {
                let name = &cursor.src()[start..start + stop];
                if let Some(control) = name.bytes().position(|b| b.is_ascii_control()) {
                    return Err(cursor.error_at(start + control, ErrorKind::ControlCharacter));
                }
                cursor.set_pos(start + stop + 1);
                Ok(Some(ParsedScope::Sheet(Cow::Borrowed(name))))
            } else {
                Ok(None)
            }
        },
    }
}

fn selector_list<'a>(cursor: &mut Cursor<'a>, nested: bool) -> Result<OneOrMany<Complex<'a>>> {
    let mut list = OneOrMany::new();
    loop {
        cursor.skip_ws();
        if list.len() >= limits::MAX_ALTERNATIVES {
            return Err(cursor.error(ErrorKind::LimitExceeded(Limit::Alternatives)));
        }
        list.push(complex(cursor, nested)?);
        cursor.skip_ws();
        if !cursor.eat(b',') {
            return Ok(list);
        }
    }
}

fn complex<'a>(cursor: &mut Cursor<'a>, nested: bool) -> Result<Complex<'a>> {
    let first = compound(cursor)?;
    let mut rest = Vec::new();
    loop {
        let save = cursor.pos();
        let had_ws = cursor.skip_ws();
        let combinator = match cursor.peek() {
            None | Some(b',') => {
                cursor.set_pos(save);
                break;
            },
            Some(b')') if nested => {
                cursor.set_pos(save);
                break;
            },
            Some(b'>') => {
                cursor.bump();
                cursor.skip_ws();
                Combinator::Child
            },
            Some(c @ (b'+' | b'~')) => {
                return Err(cursor.error(ErrorKind::UnsupportedCombinator(char::from(c))));
            },
            Some(_) if had_ws => Combinator::Descendant,
            Some(_) => return Err(cursor.unexpected(Expected::CombinatorOrEnd)),
        };
        if rest.len() + 1 >= limits::MAX_COMPOUNDS {
            return Err(cursor.error(ErrorKind::LimitExceeded(Limit::Compounds)));
        }
        rest.push((combinator, compound(cursor)?));
    }
    Ok(Complex { first, rest })
}

fn compound<'a>(cursor: &mut Cursor<'a>) -> Result<Compound<'a>> {
    let start = cursor.pos();
    let element = match cursor.peek() {
        Some(b'*') => {
            cursor.bump();
            Element::Any
        },
        Some(b) if b.is_ascii_alphanumeric() || b == b'_' => {
            let len = cursor
                .rest()
                .iter()
                .take_while(|&&b| is_ident_byte(b))
                .count();
            cursor.set_pos(start + len);
            extend_cell_range(cursor, start);
            Element::Named(Cow::Borrowed(&cursor.src()[start..cursor.pos()]))
        },
        _ => Element::Implicit,
    };
    let mut predicates = Predicates::new();
    let mut pseudos = OneOrMany::new();
    loop {
        match cursor.peek() {
            Some(b'[') => cursor.predicates(&mut predicates, pseudos.len())?,
            Some(b':') => {
                if predicates.len() + pseudos.len() >= limits::MAX_PREDICATES {
                    return Err(cursor.error(ErrorKind::LimitExceeded(Limit::Predicates)));
                }
                cursor.bump();
                pseudos.push(pseudo(cursor)?);
            },
            _ => break,
        }
    }
    if matches!(element, Element::Implicit) && predicates.is_empty() && pseudos.is_empty() {
        return Err(cursor.unexpected(Expected::Compound));
    }
    Ok(Compound {
        element,
        predicates,
        pseudos,
    })
}

/// Extends an `A1` element to `A1:C5` when a cell reference follows the colon.
fn extend_cell_range(cursor: &mut Cursor<'_>, start: usize) {
    let src = cursor.src();
    if super::axis::CellRef::parse(&src[start..cursor.pos()]).is_none()
        || cursor.peek() != Some(b':')
    {
        return;
    }
    let after = cursor.pos() + 1;
    let len = src.as_bytes()[after..cursor.end()]
        .iter()
        .take_while(|b| b.is_ascii_alphanumeric())
        .count();
    if len > 0 && super::axis::CellRef::parse(&src[after..after + len]).is_some() {
        let boundary = src.as_bytes().get(after + len).copied();
        if !matches!(boundary, Some(b) if is_ident_byte(b)) {
            cursor.set_pos(after + len);
        }
    }
}

fn pseudo<'a>(cursor: &mut Cursor<'a>) -> Result<Pseudo<'a>> {
    let start = cursor.pos();
    let len = cursor
        .rest()
        .iter()
        .take_while(|&&b| is_ident_byte(b))
        .count();
    let name = &cursor.rest()[..len];
    let call = cursor.rest().get(len) == Some(&b'(');
    let known = |keyword: &[u8]| len == keyword.len() && starts_with_ci(name, keyword);
    if known(b"contains") || known(b"has") || known(b"not") {
        if !call {
            cursor.set_pos(start + len);
            return Err(cursor.unexpected(Expected::OpenParen));
        }
        let is_contains = known(b"contains");
        let is_has = known(b"has");
        cursor.set_pos(start + len + 1);
        if is_contains {
            return contains_argument(cursor).map(Pseudo::Contains);
        }
        cursor.enter()?;
        cursor.skip_ws();
        let list = selector_list(cursor, true)?.into_vec();
        cursor.skip_ws();
        cursor.expect(b')', Expected::CloseParen)?;
        cursor.leave();
        return Ok(if is_has {
            Pseudo::Has(list)
        } else {
            Pseudo::Not(list)
        });
    }
    if !call && known(b"empty") {
        cursor.set_pos(start + len);
        return Ok(Pseudo::Empty);
    }
    if !call && known(b"no-alt") {
        cursor.set_pos(start + len);
        return Ok(Pseudo::NoAlt);
    }
    if call && len > 0 {
        return Err(cursor.error(ErrorKind::UnknownPseudoFunction));
    }
    // OfficeCLI shorthand `cell:TBD` == `cell:contains(TBD)`.
    let text_len = cursor
        .rest()
        .iter()
        .take_while(|&&b| {
            !(is_ws(b)
                || is_quote(b)
                || b.is_ascii_control()
                || matches!(b, b',' | b'>' | b'(' | b')' | b'[' | b']' | b':' | b'\\'))
        })
        .count();
    if text_len == 0 {
        return Err(cursor.unexpected(Expected::PseudoName));
    }
    cursor.set_pos(start + text_len);
    Ok(Pseudo::Contains(Cow::Borrowed(
        &cursor.src()[start..start + text_len],
    )))
}

/// Parses the argument of `:contains(`; the cursor is just past `(`.
fn contains_argument<'a>(cursor: &mut Cursor<'a>) -> Result<Cow<'a, str>> {
    cursor.skip_ws();
    let text = if matches!(cursor.peek(), Some(b) if is_quote(b)) {
        cursor.quoted()?
    } else {
        let start = cursor.pos();
        loop {
            match cursor.peek() {
                None | Some(b')') => break,
                Some(b) if is_quote(b) => return Err(cursor.error(ErrorKind::QuoteInBareValue)),
                Some(b'[' | b']' | b'(') => return Err(cursor.unexpected(Expected::CloseParen)),
                Some(b) if b.is_ascii_control() => {
                    return Err(cursor.error(ErrorKind::ControlCharacter));
                },
                Some(_) => cursor.bump(),
            }
        }
        let raw = &cursor.src()[start..cursor.pos()];
        let text = raw.trim_end_matches(|c: char| c.is_ascii() && is_ws(c as u8));
        if text.is_empty() {
            return Err(cursor.unexpected(Expected::Value));
        }
        Cow::Borrowed(text)
    };
    cursor.skip_ws();
    cursor.expect(b')', Expected::CloseParen)?;
    Ok(text)
}

impl fmt::Debug for Element<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Element { .. }")
    }
}

impl fmt::Debug for Pseudo<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Pseudo { .. }")
    }
}

impl fmt::Debug for Scope<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Scope { .. }")
    }
}
