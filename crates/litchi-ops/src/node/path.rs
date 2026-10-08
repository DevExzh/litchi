//! Node addresses and their canonical path text (record 0010 §2.6).
//!
//! A node keeps a compact address: [`PathStep`]s naming a kind or a sheet,
//! each with at most one selector (a position or a stable identity). Its
//! [`NodePath`] renders the address on demand through the selector grammar's
//! own `Step` display, one step at a time, so the text is the grammar's
//! canonical spelling (quoting included) and parses back to an equal
//! [`Path`](litchi_core::selector::Path). Rendering allocates nothing and
//! builds no `String`.

use std::borrow::Cow;
use std::fmt::{self, Write as _};

use litchi_core::Position;
use litchi_core::selector::{
    CompareOp, Condition, Expr, Key, Path, Predicate, Step, Value as Operand,
};

/// The canonical path of a node: `/`, `/body/p[3]`, `/body/p[@paraId=1A2B3C4D]`,
/// `/Sheet1/A1`, `/slide[1]/shape[@id=5]`.
///
/// A `NodePath` borrows the steps its node stores. `Display` writes the
/// grammar's canonical text without allocating; [`NodePath::to_path`] builds
/// the equal [`Path`] value.
///
/// A path deeper than the grammar's step limit (64,
/// `litchi_core::selector::limits::MAX_STEPS`) still renders, but does not
/// parse back.
///
/// ```
/// use litchi_core::Position;
/// use litchi_core::selector::Path;
/// use litchi_ops::node::{NodePath, PathStep};
/// use litchi_ops::node::raw::{Ident, identity_step};
///
/// let steps = [
///     PathStep::new("body"),
///     identity_step("p", "paraId", Ident::Text("1A2B3C4D")),
///     PathStep::at("r", Position::new(1)),
/// ];
/// let path = NodePath::new(&steps);
/// assert_eq!(path.to_string(), "/body/p[@paraId=1A2B3C4D]/r[2]");
/// assert_eq!(Path::parse(&path.to_string())?, path.to_path());
/// assert_eq!(NodePath::ROOT.to_string(), "/");
/// # Ok::<(), litchi_core::selector::ParseError>(())
/// ```
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct NodePath<'a> {
    steps: &'a [PathStep<'a>],
}

impl<'a> NodePath<'a> {
    /// The document root, `/`.
    pub const ROOT: NodePath<'static> = NodePath { steps: &[] };

    /// The path through `steps`, from the root.
    #[must_use]
    pub const fn new(steps: &'a [PathStep<'a>]) -> Self {
        Self { steps }
    }

    /// The steps, from the root.
    #[must_use]
    pub const fn steps(self) -> &'a [PathStep<'a>] {
        self.steps
    }

    /// Whether this is the root path `/`.
    #[must_use]
    pub const fn is_root(self) -> bool {
        self.steps.is_empty()
    }

    /// The equal grammar value. It allocates where a [`Path`] does: past six
    /// steps, and for each numeric identity.
    ///
    /// ```
    /// use litchi_ops::node::{NodePath, PathStep};
    /// use litchi_ops::node::raw::{Ident, identity_step};
    ///
    /// let steps = [
    ///     PathStep::at("slide", litchi_core::Position::new(0)),
    ///     identity_step("shape", "id", Ident::Number(5)),
    /// ];
    /// let path = NodePath::new(&steps).to_path();
    /// assert_eq!(path.to_string(), "/slide[1]/shape[@id=5]");
    /// ```
    #[must_use]
    /// Native identity steps require `path::raw::path` before resolution;
    /// ordinary format adapters must produce semantic names or positions.
    pub fn to_path(self) -> Path<'a> {
        Path::from_steps(self.steps.iter().map(|step| step.to_step()))
    }
}

impl fmt::Display for NodePath<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.steps.is_empty() {
            return f.write_char('/');
        }
        for step in self.steps {
            f.write_char('/')?;
            fmt::Display::fmt(step, f)?;
        }
        Ok(())
    }
}

/// One step of a node address: a name and at most one selector.
///
/// Element names come from the format's static tables; sheet names and cell
/// references are borrowed from the node. Native identity construction is
/// available only through [`super::raw::identity_step`].
///
/// ```compile_fail
/// use litchi_ops::node::PathStep;
/// let _ = PathStep::identity("shape", "id", litchi_ops::node::raw::Ident::Number(5));
/// ```
///
/// ```
/// use litchi_core::Position;
/// use litchi_ops::node::PathStep;
/// use litchi_ops::node::raw::{Ident, identity_step};
///
/// assert_eq!(PathStep::new("Sheet1").to_string(), "Sheet1");
/// assert_eq!(PathStep::at("tbl", Position::new(1)).to_string(), "tbl[2]");
/// assert_eq!(
///     identity_step("comment", "commentId", Ident::Number(3)).to_string(),
///     "comment[@commentId=3]"
/// );
/// // Names the bare grammar cannot carry are quoted.
/// assert_eq!(PathStep::new("Q1/Q2").to_string(), r#""Q1/Q2""#);
/// ```
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct PathStep<'a> {
    name: &'a str,
    select: Select<'a>,
}

impl<'a> PathStep<'a> {
    /// A step without a selector: `body`, `Sheet1`, `A1`.
    #[must_use]
    pub const fn new(name: &'a str) -> Self {
        Self {
            name,
            select: Select::Bare,
        }
    }

    /// The step `name[N]` for the zero-based `position` (written 1-based).
    #[must_use]
    pub const fn at(name: &'a str, position: Position) -> Self {
        Self {
            name,
            select: Select::Position(position),
        }
    }

    /// The step `name[@key=value]`, addressing by stable identity.
    #[must_use]
    pub(crate) const fn identity(name: &'a str, key: &'a str, value: Ident<'a>) -> Self {
        Self {
            name,
            select: Select::Identity { key, value },
        }
    }

    /// The step's grammar counterpart, when it has at most one selector of a
    /// kind an ordinary node address uses (a position). Native identities
    /// require [`super::raw::step_from_grammar`].
    ///
    /// ```
    /// use litchi_core::selector::Path;
    /// use litchi_ops::node::PathStep;
    ///
    /// let path = Path::parse("/body/p[2]")?;
    /// let step = PathStep::from_grammar(&path.steps()[1]).expect("a position step");
    /// assert_eq!(step.to_string(), "p[2]");
    /// assert!(PathStep::from_grammar(&Path::parse("/p[last()]")?.steps()[0]).is_none());
    /// # Ok::<(), litchi_core::selector::ParseError>(())
    /// ```
    #[must_use]
    pub fn from_grammar(step: &'a Step<'_>) -> Option<Self> {
        match step.predicates() {
            [] => Some(Self::new(step.name())),
            [Predicate::Index(position)] => Some(Self::at(step.name(), *position)),
            _ => None,
        }
    }

    pub(crate) fn from_raw_grammar(step: &'a Step<'_>) -> Option<Self> {
        let select = match step.predicates() {
            [] => Select::Bare,
            [Predicate::Index(position)] => Select::Position(*position),
            [predicate] => {
                let (key, value) = predicate.identity()?;
                Select::Identity {
                    key,
                    value: Ident::Text(value),
                }
            },
            _ => return None,
        };
        Some(Self {
            name: step.name(),
            select,
        })
    }

    /// The step name.
    #[must_use]
    pub const fn name(self) -> &'a str {
        self.name
    }

    /// The step's selector.
    #[must_use]
    #[cfg(test)]
    pub(crate) const fn select(self) -> Select<'a> {
        self.select
    }

    /// The equal grammar step. It allocates only for a numeric identity.
    #[must_use]
    pub fn to_step(self) -> Step<'a> {
        let step = Step::new(self.name);
        match self.select {
            Select::Bare => step,
            Select::Position(position) => step.with(Predicate::Index(position)),
            Select::Identity { key, value } => {
                let value = match value {
                    Ident::Text(text) => Cow::Borrowed(text),
                    Ident::Number(number) => Cow::Owned(number.to_string()),
                };
                step.with(identity(key, value))
            },
        }
    }
}

impl fmt::Display for PathStep<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut digits = Digits::default();
        let step = Step::new(self.name);
        let step = match self.select {
            Select::Bare => step,
            Select::Position(position) => step.with(Predicate::Index(position)),
            Select::Identity { key, value } => {
                let value = match value {
                    Ident::Text(text) => text,
                    Ident::Number(number) => digits.write(number),
                };
                step.with(identity(key, Cow::Borrowed(value)))
            },
        };
        fmt::Display::fmt(&step, f)
    }
}

/// The selector of a [`PathStep`].
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub(crate) enum Select<'a> {
    /// No selector: `body`, `Sheet1`, `A1`.
    Bare,
    /// `[N]`: the zero-based position, written 1-based.
    Position(Position),
    /// `[@key=value]`: a stable identity such as `@paraId` or `@id`.
    Identity {
        /// The identity key, without `@`.
        key: &'a str,
        /// The identity value.
        value: Ident<'a>,
    },
}

/// The value of an identity selector.
///
/// ```
/// use litchi_ops::node::raw::Ident;
///
/// assert_eq!(Ident::from(7_u32), Ident::Number(7));
/// assert_eq!(Ident::from("1A2B3C4D"), Ident::Text("1A2B3C4D"));
/// ```
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub enum Ident<'a> {
    /// Text, such as a `w14:paraId` or a shape name.
    Text(&'a str),
    /// A number, such as a `cNvPr` id or a comment id; written in decimal.
    Number(u64),
}

impl<'a> From<&'a str> for Ident<'a> {
    fn from(text: &'a str) -> Self {
        Self::Text(text)
    }
}

impl From<u64> for Ident<'_> {
    fn from(number: u64) -> Self {
        Self::Number(number)
    }
}

impl From<u32> for Ident<'_> {
    fn from(number: u32) -> Self {
        Self::Number(number.into())
    }
}

/// The predicate `[@key=value]`.
fn identity<'a>(key: &'a str, value: Cow<'a, str>) -> Predicate<'a> {
    Predicate::Filter(Expr::Condition(Condition::Compare {
        key: Key::attribute(key),
        op: CompareOp::Eq,
        value: Operand::Text(value),
    }))
}

/// A stack buffer for the decimal text of a `u64`.
#[derive(Default)]
struct Digits {
    bytes: [u8; 20],
}

impl Digits {
    fn write(&mut self, mut number: u64) -> &str {
        let mut start = self.bytes.len();
        loop {
            start -= 1;
            // `number % 10` is a single digit.
            self.bytes[start] = b'0' + (number % 10) as u8;
            number /= 10;
            if number == 0 {
                break;
            }
        }
        // ASCII digits are always UTF-8.
        std::str::from_utf8(&self.bytes[start..]).unwrap_or_default()
    }
}

impl fmt::Debug for NodePath<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("NodePath").finish_non_exhaustive()
    }
}

impl fmt::Debug for PathStep<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("PathStep").finish_non_exhaustive()
    }
}

impl fmt::Debug for Select<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Select").finish_non_exhaustive()
    }
}

impl fmt::Debug for Ident<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Ident").finish_non_exhaustive()
    }
}
