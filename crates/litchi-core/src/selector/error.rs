//! Typed parse errors for the path and selector grammar.

use std::fmt;

use super::limits;

/// Which grammar an input was parsed against.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Grammar {
    /// An element path such as `/body/p[3]`.
    Path,
    /// A CSS-like query selector such as `paragraph[style=Heading1]`.
    Query,
}

/// The stable, machine-readable error code of the unified operations layer.
///
/// The parser produces only [`ErrorCode::InvalidPath`],
/// [`ErrorCode::InvalidSelector`], and [`ErrorCode::MaxDepthExceeded`].
/// [`ErrorCode::PathNotFound`] is reserved here so that resolution-time misses
/// share the same vocabulary: a syntactically valid path that names nothing is
/// a different failure from a malformed one.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ErrorCode {
    /// The path string is not valid path syntax.
    InvalidPath,
    /// The query string is not valid selector syntax.
    InvalidSelector,
    /// A nesting or depth limit was exceeded.
    MaxDepthExceeded,
    /// A valid path did not resolve to an element. Resolution-time only; the
    /// parser never returns it.
    PathNotFound,
}

impl ErrorCode {
    /// Returns the snake-case code string (`invalid_path`, ...).
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::InvalidPath => "invalid_path",
            Self::InvalidSelector => "invalid_selector",
            Self::MaxDepthExceeded => "max_depth_exceeded",
            Self::PathNotFound => "path_not_found",
        }
    }
}

impl fmt::Display for ErrorCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

/// A bounded parser dimension (see [`limits`]).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Limit {
    /// [`limits::MAX_INPUT_BYTES`].
    InputBytes,
    /// [`limits::MAX_STEPS`].
    Steps,
    /// [`limits::MAX_PREDICATES`].
    Predicates,
    /// [`limits::MAX_NESTING`].
    Nesting,
    /// [`limits::MAX_TERMS`].
    Terms,
    /// [`limits::MAX_ALTERNATIVES`].
    Alternatives,
    /// [`limits::MAX_COMPOUNDS`].
    Compounds,
}

impl Limit {
    /// Returns the configured maximum for this dimension.
    #[must_use]
    pub const fn max(self) -> usize {
        match self {
            Self::InputBytes => limits::MAX_INPUT_BYTES,
            Self::Steps => limits::MAX_STEPS,
            Self::Predicates => limits::MAX_PREDICATES,
            Self::Nesting => limits::MAX_NESTING,
            Self::Terms => limits::MAX_TERMS,
            Self::Alternatives => limits::MAX_ALTERNATIVES,
            Self::Compounds => limits::MAX_COMPOUNDS,
        }
    }

    const fn noun(self) -> &'static str {
        match self {
            Self::InputBytes => "input bytes",
            Self::Steps => "path steps",
            Self::Predicates => "predicates on one step",
            Self::Nesting => "nesting levels",
            Self::Terms => "conditions in one bracket",
            Self::Alternatives => "comma-separated alternatives",
            Self::Compounds => "chained compounds",
        }
    }
}

/// What the parser found at the error offset.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub enum Found {
    /// A character.
    Char(char),
    /// The end of the input.
    End,
}

impl fmt::Display for Found {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Char(_) => f.write_str("a character"),
            Self::End => f.write_str("end of input"),
        }
    }
}

/// The token class the parser expected at the error offset.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Expected {
    /// A leading `/`.
    Root,
    /// A path segment name.
    SegmentName,
    /// `/`, `[`, or the end of the path after a step.
    StepEnd,
    /// The closing `]` of a predicate.
    CloseBracket,
    /// A predicate body: `N`, `last()`, `key`, or `key op value`.
    Predicate,
    /// A comparison operator.
    Operator,
    /// A value after an operator.
    Value,
    /// An opening `(`.
    OpenParen,
    /// A closing `)`.
    CloseParen,
    /// A selector compound: element type, `*`, `[...]`, or `:pseudo`.
    Compound,
    /// A combinator (`>` or whitespace), `,`, or the end of the selector.
    CombinatorOrEnd,
    /// The `!` after an Excel-style sheet scope.
    SheetBang,
    /// A pseudo-class name after `:`.
    PseudoName,
    /// A closing quote.
    CloseQuote,
}

impl fmt::Display for Expected {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::Root => "'/'",
            Self::SegmentName => "a path segment name",
            Self::StepEnd => "'/', '[', or end of path",
            Self::CloseBracket => "']'",
            Self::Predicate => "a predicate (N, last(), key, or key op value)",
            Self::Operator => "an operator (=, !=, ~=, >, <, >=, <=), ']' or ')'",
            Self::Value => "a value",
            Self::OpenParen => "'('",
            Self::CloseParen => "')'",
            Self::Compound => "an element type, '*', '[', or ':'",
            Self::CombinatorOrEnd => "'>', whitespace, ',', or end of selector",
            Self::SheetBang => "'!' after the sheet name",
            Self::PseudoName => "a pseudo-class name",
            Self::CloseQuote => "a closing quote",
        })
    }
}

/// The specific syntax failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ErrorKind {
    /// The input was empty or whitespace only.
    Empty,
    /// A token other than the expected one was found.
    Unexpected {
        /// The token class the parser required.
        expected: Expected,
        /// What was actually there.
        found: Found,
    },
    /// A path contained `//` or ended with `/`.
    EmptySegment,
    /// A 1-based index was `0`.
    IndexZero,
    /// A 1-based index was negative.
    IndexNegative,
    /// A 1-based index does not fit the platform's position type.
    IndexOverflow,
    /// A quoted token was not closed.
    UnterminatedQuote,
    /// An unquoted name contained an ASCII control character.
    ControlCharacter,
    /// An unquoted value contained a quote character.
    QuoteInBareValue,
    /// An unquoted value contained `*`; wildcards are not supported.
    Wildcard,
    /// An unquoted value began with `=`, `~`, or `!`, which indicates a
    /// mistyped operator such as `==`.
    MisparsedOperator,
    /// A regex value `r"..."` was used with an operator other than `~=`.
    RegexRequiresContains,
    /// A sibling combinator (`+` or `~`) was used; only `>` and descendant
    /// whitespace are supported.
    UnsupportedCombinator(char),
    /// A functional pseudo-class other than `:contains()`, `:has()`, or
    /// `:not()` was used.
    UnknownPseudoFunction,
    /// A bound from [`limits`] was exceeded.
    LimitExceeded(Limit),
}

impl fmt::Display for ErrorKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => f.write_str("input is empty"),
            Self::Unexpected { expected, found } => {
                write!(f, "expected {expected}, found {found}")
            },
            Self::EmptySegment => {
                f.write_str("empty path segment (check for '//' or a trailing '/')")
            },
            Self::IndexZero => f.write_str("index 0 is invalid; indices are 1-based"),
            Self::IndexNegative => f.write_str("negative indices are not supported; use last()"),
            Self::IndexOverflow => f.write_str("index is too large"),
            Self::UnterminatedQuote => f.write_str("unterminated quoted token"),
            Self::ControlCharacter => f.write_str("control character in unquoted name"),
            Self::QuoteInBareValue => f.write_str("quote character inside an unquoted value"),
            Self::Wildcard => f.write_str("wildcards (*) are not supported in values"),
            Self::MisparsedOperator => {
                f.write_str("value starts with '=', '~' or '!'; the operator is malformed")
            },
            Self::RegexRequiresContains => {
                f.write_str("regex values r\"...\" are only supported with the ~= operator")
            },
            Self::UnsupportedCombinator(c) => {
                write!(
                    f,
                    "combinator {c:?} is not supported; use '>' or whitespace"
                )
            },
            Self::UnknownPseudoFunction => f.write_str(
                "unknown functional pseudo-class; supported: :contains(), :has(), :not()",
            ),
            Self::LimitExceeded(limit) => {
                write!(f, "more than {} {}", limit.max(), limit.noun())
            },
        }
    }
}

/// A typed path or selector syntax error with the byte offset of the failure.
///
/// ```
/// use litchi_core::selector::{ErrorCode, ErrorKind, Path};
///
/// let error = Path::parse("/body/p[0]").unwrap_err();
/// assert_eq!(error.kind(), ErrorKind::IndexZero);
/// assert_eq!(error.offset(), 8);
/// assert_eq!(error.code(), ErrorCode::InvalidPath);
/// assert_eq!(error.code().as_str(), "invalid_path");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, thiserror::Error)]
#[error("{} at byte {offset}: {kind}", grammar_noun(*.grammar))]
pub struct ParseError {
    grammar: Grammar,
    offset: usize,
    kind: ErrorKind,
}

const fn grammar_noun(grammar: Grammar) -> &'static str {
    match grammar {
        Grammar::Path => "invalid path",
        Grammar::Query => "invalid selector",
    }
}

impl ParseError {
    pub(crate) const fn new(grammar: Grammar, offset: usize, kind: ErrorKind) -> Self {
        Self {
            grammar,
            offset,
            kind,
        }
    }

    /// Returns the grammar that rejected the input.
    #[must_use]
    pub const fn grammar(&self) -> Grammar {
        self.grammar
    }

    /// Returns the byte offset into the original input where parsing failed.
    #[must_use]
    pub const fn offset(&self) -> usize {
        self.offset
    }

    /// Returns the error with its offset moved `base` bytes later.
    ///
    /// A caller that parses a suffix of a larger text, such as the part after
    /// `!` in the spreadsheet form `Sheet1!A1`, uses this to report the offset
    /// in the text the user wrote. The offset saturates at `usize::MAX`.
    ///
    /// ```
    /// use litchi_core::selector::{ErrorKind, Path};
    ///
    /// let text = "Sheet1!A1[0]";
    /// let error = Path::parse_relative(&text[7..]).unwrap_err().offset_by(7);
    /// assert_eq!((error.kind(), error.offset()), (ErrorKind::IndexZero, 10));
    /// ```
    #[must_use]
    pub const fn offset_by(self, base: usize) -> Self {
        Self {
            offset: self.offset.saturating_add(base),
            ..self
        }
    }

    /// Returns the specific failure.
    #[must_use]
    pub const fn kind(&self) -> ErrorKind {
        self.kind
    }

    /// Returns the stable error code: `invalid_path` / `invalid_selector` for
    /// syntax errors, `max_depth_exceeded` for step and nesting limits.
    #[must_use]
    pub const fn code(&self) -> ErrorCode {
        match self.kind {
            ErrorKind::LimitExceeded(Limit::Steps | Limit::Nesting) => ErrorCode::MaxDepthExceeded,
            _ => match self.grammar {
                Grammar::Path => ErrorCode::InvalidPath,
                Grammar::Query => ErrorCode::InvalidSelector,
            },
        }
    }

    /// Returns a short corrective hint for failures with an obvious fix.
    #[must_use]
    pub const fn suggestion(&self) -> Option<&'static str> {
        match self.kind {
            ErrorKind::Wildcard => Some("use ~= for a substring match, e.g. [name~=Excel]"),
            ErrorKind::MisparsedOperator => Some("supported operators are =, !=, ~=, >, <, >=, <="),
            ErrorKind::RegexRequiresContains => {
                Some("write the regex with ~=, e.g. [text~=r\"^Q[1-4]\"]")
            },
            ErrorKind::IndexZero | ErrorKind::IndexNegative => {
                Some("indices are 1-based; use [1] for the first element and [last()] for the last")
            },
            ErrorKind::EmptySegment => Some("remove the doubled or trailing '/'"),
            ErrorKind::QuoteInBareValue | ErrorKind::ControlCharacter => {
                Some("quote the value with \"...\" and escape inner quotes with a backslash")
            },
            ErrorKind::UnsupportedCombinator(_) => {
                Some("use 'a > b' for children or 'a b' for descendants")
            },
            _ => None,
        }
    }
}

impl fmt::Debug for Found {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(self, f)
    }
}
