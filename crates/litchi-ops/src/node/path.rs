//! Node addresses and their canonical path text (record 0010 §2.6).
//!
//! A node keeps a compact address: [`PathStep`]s naming a kind or a sheet,
//! each with at most one selector (a position or a stable identity), or a
//! worksheet cell or range written in A1 notation. Its [`NodePath`] renders
//! the address on demand through the selector grammar's own `Step` display,
//! one step at a time, so the text is the grammar's canonical spelling
//! (quoting included) and parses back to an equal
//! [`Path`](litchi_core::selector::Path). Rendering allocates nothing and
//! builds no `String`.

use std::borrow::Cow;
use std::fmt::{self, Write as _};

use litchi_core::Position;
use litchi_core::selector::{
    CompareOp, Condition, Expr, Key, Path, Predicate, Step, Value as Operand,
};

/// A node address: ordinary adapters emit positional paths such as `/body/p[3]`.
/// Explicit raw construction can also represent `/body/p[@paraId=1A2B3C4D]`,
/// `/Sheet1/A1`, `/slide[1]/shape[@id=5]`.
///
/// A `NodePath` borrows the steps its node stores, or carries up to four
/// steps by value: a node whose address is short and cheap to spell out (a
/// worksheet cell, a slide's shape) writes its steps on demand instead of
/// storing them, which keeps it small to move (record 0262). `Display`
/// writes the grammar's canonical text without allocating;
/// [`NodePath::to_path`] builds the equal [`Path`] value.
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
#[derive(Clone, Copy)]
pub struct NodePath<'a> {
    repr: Repr<'a>,
}

/// The most steps a [`NodePath`] carries by value.
pub(crate) const INLINE_STEPS: usize = 4;

#[derive(Clone, Copy)]
#[allow(
    clippy::large_enum_variant,
    reason = "a path is a short-lived value; carrying a few steps inline is what spares a node \
              from storing them"
)]
enum Repr<'a> {
    /// Steps the node stores.
    Borrowed(&'a [PathStep<'a>]),
    /// Steps written for this path, the first `len` of `steps`.
    Inline {
        steps: [PathStep<'a>; INLINE_STEPS],
        len: u8,
    },
}

/// The step that fills an unused inline slot.
const EMPTY_STEP: PathStep<'static> = PathStep::new("");

impl<'a> NodePath<'a> {
    /// The document root, `/`.
    pub const ROOT: NodePath<'static> = NodePath {
        repr: Repr::Borrowed(&[]),
    };

    /// The path through `steps`, from the root.
    #[must_use]
    pub const fn new(steps: &'a [PathStep<'a>]) -> Self {
        Self {
            repr: Repr::Borrowed(steps),
        }
    }

    /// The path through `steps`, from the root, carried by value; `None`
    /// for more than four steps.
    #[cfg_attr(
        not(feature = "pptx"),
        allow(dead_code, reason = "the PPTX adapter writes its short paths this way")
    )]
    pub(crate) fn inline(steps: &[PathStep<'a>]) -> Option<Self> {
        let len = u8::try_from(steps.len())
            .ok()
            .filter(|&len| usize::from(len) <= INLINE_STEPS)?;
        let mut inline = [EMPTY_STEP; INLINE_STEPS];
        inline[..steps.len()].copy_from_slice(steps);
        Some(Self {
            repr: Repr::Inline { steps: inline, len },
        })
    }

    /// The path of `len` steps from the root, each written by `step` from
    /// its index, carried by value; `None` for more than four steps. The
    /// steps are written in place (record 0287), with no staging array.
    #[cfg_attr(
        not(any(feature = "doc", feature = "rtf", feature = "odt")),
        allow(
            dead_code,
            reason = "the DOC, RTF and ODT adapters spell their compact paths this way"
        )
    )]
    pub(crate) fn inline_with(
        len: usize,
        mut step: impl FnMut(usize) -> PathStep<'a>,
    ) -> Option<Self> {
        let len = u8::try_from(len)
            .ok()
            .filter(|&len| usize::from(len) <= INLINE_STEPS)?;
        let mut steps = [EMPTY_STEP; INLINE_STEPS];
        for (index, slot) in steps[..usize::from(len)].iter_mut().enumerate() {
            *slot = step(index);
        }
        Some(Self {
            repr: Repr::Inline { steps, len },
        })
    }

    /// The path of `first` and, when given, `second`, carried by value.
    #[cfg_attr(
        not(any(feature = "xlsx", feature = "xls", feature = "ods")),
        allow(
            dead_code,
            reason = "the XLSX, XLS and ODS adapters write their paths this way"
        )
    )]
    pub(crate) fn short(first: PathStep<'a>, second: Option<PathStep<'a>>) -> Self {
        let mut steps = [EMPTY_STEP; INLINE_STEPS];
        steps[0] = first;
        let len = match second {
            Some(second) => {
                steps[1] = second;
                2
            },
            None => 1,
        };
        Self {
            repr: Repr::Inline { steps, len },
        }
    }

    /// The steps, from the root.
    #[must_use]
    pub fn steps(&self) -> &[PathStep<'a>] {
        match &self.repr {
            Repr::Borrowed(steps) => steps,
            Repr::Inline { steps, len } => &steps[..usize::from(*len)],
        }
    }

    /// Whether this is the root path `/`.
    #[must_use]
    pub fn is_root(&self) -> bool {
        self.steps().is_empty()
    }

    /// The equal grammar value. It allocates where a [`Path`] does: past six
    /// steps, for numeric identities, and for A1 cell, range or column text.
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
        Path::from_steps(self.steps().iter().map(|step| step.to_step()))
    }
}

/// Paths are equal when their steps are, however they are carried.
impl PartialEq for NodePath<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.steps() == other.steps()
    }
}

impl Eq for NodePath<'_> {}

impl std::hash::Hash for NodePath<'_> {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.steps().hash(state);
    }
}

impl fmt::Display for NodePath<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let steps = self.steps();
        if steps.is_empty() {
            return f.write_char('/');
        }
        for step in steps {
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

    /// The worksheet cell step `B3`, written in A1 notation from its
    /// zero-based `column` and `row`.
    #[must_use]
    #[cfg(any(
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "ods",
        feature = "numbers"
    ))]
    pub(crate) const fn cell(column: u32, row: u32) -> Self {
        Self {
            name: "",
            select: Select::Cell { column, row },
        }
    }

    /// The worksheet range step `A1:C5` from its first and last cells, each
    /// a zero-based `(column, row)`.
    #[must_use]
    #[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
    pub(crate) const fn range(first: (u32, u32), last: (u32, u32)) -> Self {
        Self {
            name: "",
            select: Select::Range { first, last },
        }
    }

    /// The step `name[B]`: a worksheet column by its letters, for the
    /// zero-based `column`.
    #[must_use]
    #[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
    pub(crate) const fn column(name: &'a str, column: u32) -> Self {
        Self {
            name,
            select: Select::Column(column),
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

    /// The equal grammar step. It allocates for numeric identities and A1
    /// cell, range or column text.
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
            Select::Cell { .. } | Select::Range { .. } => Step::new(Cow::Owned(self.to_string())),
            Select::Column(column) => {
                let mut letters = A1::default();
                let letters = letters.column(column).to_owned();
                step.with(bare(Cow::Owned(letters)))
            },
        }
    }
}

impl fmt::Display for PathStep<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut digits = Digits::default();
        let mut a1 = A1::default();
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
            // A1 text is always a plain name in the grammar: letters, digits
            // and one colon.
            Select::Cell { column, row } => return f.write_str(a1.cell(column, row)),
            Select::Range { first, last } => return f.write_str(a1.range(first, last)),
            Select::Column(column) => step.with(bare(Cow::Borrowed(a1.column(column)))),
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
    /// A worksheet cell, written `B3`: the whole step, with no name.
    #[cfg_attr(
        not(feature = "xlsx"),
        allow(dead_code, reason = "constructed only by the XLSX adapter")
    )]
    Cell {
        /// The zero-based column.
        column: u32,
        /// The zero-based row.
        row: u32,
    },
    /// A worksheet range, written `A1:C5`: the whole step, with no name.
    #[cfg_attr(
        not(feature = "xlsx"),
        allow(dead_code, reason = "constructed only by the XLSX adapter")
    )]
    Range {
        /// The first cell: zero-based `(column, row)`.
        first: (u32, u32),
        /// The last cell: zero-based `(column, row)`.
        last: (u32, u32),
    },
    /// `[B]`: a worksheet column by its letters, for the zero-based column.
    #[cfg_attr(
        not(feature = "xlsx"),
        allow(dead_code, reason = "constructed only by the XLSX adapter")
    )]
    Column(u32),
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

/// The bare-token predicate `[B]`.
fn bare(token: Cow<'_, str>) -> Predicate<'_> {
    Predicate::Filter(Expr::Condition(Condition::Exists(Key::new(token))))
}

/// A stack buffer for A1 text: column letters, a cell (`XFD1048576`) or a
/// range of two cells.
#[derive(Default)]
struct A1 {
    bytes: [u8; 24],
    len: usize,
}

impl A1 {
    /// The letters of the zero-based `column`: `A` for 0.
    fn column(&mut self, column: u32) -> &str {
        self.len = 0;
        self.push_column(column);
        self.text()
    }

    /// `B3` for the cell at zero-based column 1, row 2.
    fn cell(&mut self, column: u32, row: u32) -> &str {
        self.len = 0;
        self.push_cell(column, row);
        self.text()
    }

    /// `A1:C5` for zero-based `(column, row)` corners, as given; a one-cell
    /// range writes one cell.
    fn range(&mut self, first: (u32, u32), last: (u32, u32)) -> &str {
        self.len = 0;
        self.push_cell(first.0, first.1);
        if last != first {
            self.push(b':');
            self.push_cell(last.0, last.1);
        }
        self.text()
    }

    fn push(&mut self, byte: u8) {
        if let Some(slot) = self.bytes.get_mut(self.len) {
            *slot = byte;
            self.len += 1;
        }
    }

    fn push_column(&mut self, column: u32) {
        let mut letters = [0_u8; 7];
        let mut start = letters.len();
        let mut number = u64::from(column) + 1;
        while number > 0 && start > 0 {
            start -= 1;
            letters[start] = b'A' + ((number - 1) % 26) as u8;
            number = (number - 1) / 26;
        }
        for &letter in &letters[start..] {
            self.push(letter);
        }
    }

    fn push_cell(&mut self, column: u32, row: u32) {
        self.push_column(column);
        let mut digits = Digits::default();
        for &digit in digits.write(u64::from(row) + 1).as_bytes() {
            self.push(digit);
        }
    }

    fn text(&self) -> &str {
        // Letters, digits and a colon are always UTF-8.
        std::str::from_utf8(&self.bytes[..self.len]).unwrap_or_default()
    }
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
