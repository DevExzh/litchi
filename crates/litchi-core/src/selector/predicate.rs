//! Bracket predicates: positional indices and boolean attribute filters.

use std::borrow::Cow;
use std::fmt::{self, Display as _, Write as _};

use super::Position;
use super::store::OneOrMany;
use super::text::{is_plain_key, is_plain_value, write_quoted};

/// The predicates attached to one path step or selector compound. One inline
/// slot covers the common `p[3]` / `shape[@name=Title]` forms without a heap
/// allocation.
pub(super) type Predicates<'a> = OneOrMany<Predicate<'a>>;

/// One `[...]` predicate.
///
/// ```
/// use litchi_core::selector::{CompareOp, Condition, Expr, Path, Position, Predicate};
///
/// let path = Path::parse("/slide[last()]/shape[@name=Title]")?;
/// assert_eq!(path.steps()[0].predicates(), &[Predicate::Last]);
/// let Predicate::Filter(Expr::Condition(Condition::Compare { key, op, value })) =
///     &path.steps()[1].predicates()[0]
/// else {
///     unreachable!()
/// };
/// assert!(key.is_attribute());
/// assert_eq!((key.name(), *op, value.as_str()), ("name", CompareOp::Eq, "Title"));
/// # Ok::<(), litchi_core::selector::ParseError>(())
/// ```
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Predicate<'a> {
    /// `[N]`: the string form is 1-based; the stored [`Position`] is
    /// zero-based. `[0]`, negative, and overflowing indices are rejected at
    /// parse time. `Position::new(usize::MAX)` has no 1-based string form.
    Index(Position),
    /// `[last()]`: the last matching sibling, resolved at lookup time.
    Last,
    /// `[expr]`: an attribute existence test, comparison, or boolean
    /// combination.
    Filter(Expr<'a>),
}

impl<'a> Predicate<'a> {
    /// Returns `true` for addressing predicates that select by position or
    /// stable identity rather than by content: `[N]`, `[last()]`, a bare
    /// token such as `col[B]` or `formfield[agree_terms]`, and a single
    /// `[@key=value]` equality. This mirrors OfficeCLI's split between a
    /// navigable path and a content-filter path.
    #[must_use]
    pub fn is_structural(&self) -> bool {
        match self {
            Self::Index(_) | Self::Last => true,
            Self::Filter(Expr::Condition(Condition::Exists(_))) => true,
            Self::Filter(Expr::Condition(Condition::Compare { key, op, .. })) => {
                key.is_attribute() && *op == CompareOp::Eq
            },
            Self::Filter(_) => false,
        }
    }

    /// Returns the token of a bare-key predicate such as `col[B]`,
    /// `placeholder[ctrTitle]`, or `cell[formula]`.
    #[must_use]
    pub fn bare_key(&self) -> Option<&str> {
        match self {
            Self::Filter(Expr::Condition(Condition::Exists(key))) if !key.is_attribute() => {
                Some(key.name())
            },
            _ => None,
        }
    }

    /// Returns `(name, value)` of a stable-identity predicate
    /// `[@name=value]` (`@paraId`, `@name`, `@id`, ...).
    #[must_use]
    pub fn identity(&self) -> Option<(&str, &str)> {
        match self {
            Self::Filter(Expr::Condition(Condition::Compare {
                key,
                op: CompareOp::Eq,
                value: Value::Text(value),
            })) if key.is_attribute() => Some((key.name(), value)),
            _ => None,
        }
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Predicate<'static> {
        match self {
            Self::Index(position) => Predicate::Index(position),
            Self::Last => Predicate::Last,
            Self::Filter(expr) => Predicate::Filter(expr.into_owned()),
        }
    }
}

impl fmt::Display for Predicate<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_char('[')?;
        match self {
            // u128 keeps `usize::MAX + 1` printable instead of panicking.
            Self::Index(position) => write!(f, "{}", position.get() as u128 + 1)?,
            Self::Last => f.write_str("last()")?,
            Self::Filter(expr) => expr.fmt(f)?,
        }
        f.write_char(']')
    }
}

/// A boolean filter expression inside one bracket.
///
/// `and` binds tighter than `or`; `not(...)` and parentheses group. Several
/// stacked brackets (`[a][b]`) are separate predicates, combined with `and`
/// by resolution. `And` and `Or` built through the parser always hold at least
/// two operands; display of hand-built single-operand forms is not guaranteed
/// to round-trip.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Expr<'a> {
    /// A single condition.
    Condition(Condition<'a>),
    /// All operands hold.
    And(Vec<Expr<'a>>),
    /// Any operand holds.
    Or(Vec<Expr<'a>>),
    /// The operand does not hold.
    Not(Box<Expr<'a>>),
}

impl<'a> Expr<'a> {
    /// Calls `visit` for every leaf condition in document order.
    pub fn for_each_condition(&self, visit: &mut impl FnMut(&Condition<'a>)) {
        match self {
            Self::Condition(condition) => visit(condition),
            Self::And(parts) | Self::Or(parts) => {
                for part in parts {
                    part.for_each_condition(visit);
                }
            },
            Self::Not(inner) => inner.for_each_condition(visit),
        }
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Expr<'static> {
        match self {
            Self::Condition(condition) => Expr::Condition(condition.into_owned()),
            Self::And(parts) => Expr::And(parts.into_iter().map(Expr::into_owned).collect()),
            Self::Or(parts) => Expr::Or(parts.into_iter().map(Expr::into_owned).collect()),
            Self::Not(inner) => Expr::Not(Box::new(inner.into_owned())),
        }
    }
}

impl fmt::Display for Expr<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Condition(condition) => condition.fmt(f),
            Self::And(parts) => write_joined(f, parts, " and ", |part| {
                matches!(part, Expr::And(_) | Expr::Or(_))
            }),
            Self::Or(parts) => write_joined(f, parts, " or ", |part| matches!(part, Expr::Or(_))),
            Self::Not(inner) => write!(f, "not({inner})"),
        }
    }
}

fn write_joined(
    f: &mut fmt::Formatter<'_>,
    parts: &[Expr<'_>],
    separator: &str,
    needs_parens: impl Fn(&Expr<'_>) -> bool,
) -> fmt::Result {
    for (index, part) in parts.iter().enumerate() {
        if index > 0 {
            f.write_str(separator)?;
        }
        if needs_parens(part) {
            write!(f, "({part})")?;
        } else {
            part.fmt(f)?;
        }
    }
    Ok(())
}

/// A leaf filter condition.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Condition<'a> {
    /// `[key]`: the key is present (or, in a path, a bare addressing token
    /// such as the column letter in `col[B]`).
    Exists(Key<'a>),
    /// `[key op value]`.
    Compare {
        /// The attribute, property, or header-column name.
        key: Key<'a>,
        /// The comparison operator.
        op: CompareOp,
        /// The comparison operand.
        value: Value<'a>,
    },
}

impl Condition<'_> {
    /// Returns the condition's key.
    #[must_use]
    pub fn key(&self) -> &Key<'_> {
        match self {
            Self::Exists(key) | Self::Compare { key, .. } => key,
        }
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Condition<'static> {
        match self {
            Self::Exists(key) => Condition::Exists(key.into_owned()),
            Self::Compare { key, op, value } => Condition::Compare {
                key: key.into_owned(),
                op,
                value: value.into_owned(),
            },
        }
    }
}

impl fmt::Display for Condition<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Exists(key) => key.fmt(f),
            Self::Compare { key, op, value } => write!(f, "{key}{op}{value}"),
        }
    }
}

/// A predicate key: `name`, `@name` (an attribute locator), or a quoted
/// header name such as `"Full Name"`.
#[derive(Clone, PartialEq, Eq, Hash)]
pub struct Key<'a> {
    name: Cow<'a, str>,
    attribute: bool,
}

impl<'a> Key<'a> {
    /// Creates a plain key (`style`, `Salary`).
    pub fn new(name: impl Into<Cow<'a, str>>) -> Self {
        Self::from_parts(name.into(), false)
    }

    /// Creates an `@`-prefixed attribute key (`@paraId`, `@name`).
    pub fn attribute(name: impl Into<Cow<'a, str>>) -> Self {
        Self::from_parts(name.into(), true)
    }

    pub(super) const fn from_parts(name: Cow<'a, str>, attribute: bool) -> Self {
        Self { name, attribute }
    }

    /// Returns the key text without the `@` prefix or quotes.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns `true` for the `@name` form.
    #[must_use]
    pub const fn is_attribute(&self) -> bool {
        self.attribute
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Key<'static> {
        Key {
            name: Cow::Owned(self.name.into_owned()),
            attribute: self.attribute,
        }
    }
}

impl fmt::Display for Key<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.attribute {
            f.write_char('@')?;
        }
        if is_plain_key(&self.name) {
            f.write_str(&self.name)
        } else {
            write_quoted(f, &self.name)
        }
    }
}

/// A comparison operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CompareOp {
    /// `=`
    Eq,
    /// `!=` (the shell-escaped `\!=` is accepted on input)
    Ne,
    /// `~=` substring (or regex with an `r"..."` value)
    Contains,
    /// `>`
    Gt,
    /// `<`
    Lt,
    /// `>=`
    Ge,
    /// `<=`
    Le,
}

impl CompareOp {
    /// Returns the operator token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Eq => "=",
            Self::Ne => "!=",
            Self::Contains => "~=",
            Self::Gt => ">",
            Self::Lt => "<",
            Self::Ge => ">=",
            Self::Le => "<=",
        }
    }
}

impl fmt::Display for CompareOp {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

/// A comparison operand.
///
/// Bare and quoted spellings of the same text produce the same `Text` value;
/// quoting only changes which characters can be written.
#[derive(Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Value<'a> {
    /// Literal text, already unescaped.
    Text(Cow<'a, str>),
    /// An `r"..."` regular-expression pattern (only with `~=`). The pattern
    /// keeps its backslashes verbatim. It is stored, not compiled: the
    /// grammar layer carries no regex engine.
    Regex(Cow<'a, str>),
}

impl<'a> Value<'a> {
    /// Creates a literal text value.
    pub fn text(text: impl Into<Cow<'a, str>>) -> Self {
        Self::Text(text.into())
    }

    /// Creates a regex value when `pattern` has a quoted form the parser reads
    /// back unchanged: it must not end inside an escape and must not contain
    /// both an unescaped `"` and an unescaped `'`.
    pub fn regex(pattern: impl Into<Cow<'a, str>>) -> Option<Self> {
        let pattern = pattern.into();
        let (ok, _) = regex_quote(&pattern);
        ok.then_some(Self::Regex(pattern))
    }

    /// Returns the text or pattern.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Text(text) | Self::Regex(text) => text,
        }
    }

    /// Converts borrowed text into owned text.
    #[must_use]
    pub fn into_owned(self) -> Value<'static> {
        match self {
            Self::Text(text) => Value::Text(Cow::Owned(text.into_owned())),
            Self::Regex(text) => Value::Regex(Cow::Owned(text.into_owned())),
        }
    }
}

/// Returns whether a regex pattern is representable and which quote to use.
fn regex_quote(pattern: &str) -> (bool, char) {
    let mut unescaped_double = false;
    let mut unescaped_single = false;
    let mut escaped = false;
    for c in pattern.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        match c {
            '\\' => escaped = true,
            '"' => unescaped_double = true,
            '\'' => unescaped_single = true,
            _ => {},
        }
    }
    let quote = if unescaped_double { '\'' } else { '"' };
    (!(escaped || unescaped_double && unescaped_single), quote)
}

impl fmt::Display for Value<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Text(text) if is_plain_value(text) => f.write_str(text),
            Self::Text(text) => write_quoted(f, text),
            Self::Regex(pattern) => {
                let (_, quote) = regex_quote(pattern);
                write!(f, "r{quote}{pattern}{quote}")
            },
        }
    }
}

impl fmt::Debug for Key<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Key { .. }")
    }
}

impl fmt::Debug for Value<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Value { .. }")
    }
}
