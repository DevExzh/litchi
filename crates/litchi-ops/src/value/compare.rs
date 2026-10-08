//! Query comparison: one `[key op value]` condition compiled once and
//! evaluated against typed values without allocating.
//!
//! The semantics are OfficeCLI's `AttributeFilter.MatchOne`
//! (`OC/Core/AttributeFilter.cs`), applied to the canonical text of the
//! node's typed value, plus typed equality: a literal that coerces to the
//! key's type equals a value it coerces to (`[fill=red]` matches `#FF0000`).

use std::cmp::Ordering;
use std::fmt::Write as _;

use litchi_core::selector::{CompareOp, Condition, Value};

use crate::error::{Code, Error, Result};

use super::length::try_emu;
use super::number::parse_numeric_text;
use super::text::{StackStr, contains_ignore_case, ends_with_ci, eq_ignore_case, strip_suffix_ci};
use super::ty::{Scalar, ValueType};

/// The canonical text of every typed value fits in this many bytes: the
/// longest is a color with eight raw transforms (about 160 bytes).
const DISPLAY_CAPACITY: usize = 256;

/// OfficeCLI's tolerance for unit-qualified length equality, in EMU.
const DIMENSION_TOLERANCE: i128 = 500;

/// The operator of a compiled [`Test`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Op {
    Exists,
    Eq,
    Ne,
    Contains,
    Gt,
    Lt,
    Ge,
    Le,
    /// An operator this version does not know; it never matches.
    Never,
}

/// A compile-time warning about a condition. Warnings are not errors, as in
/// OfficeCLI; a warned condition simply never matches.
///
/// ```
/// use litchi_core::selector::CompareOp;
/// use litchi_ops::value::{Test, Warning};
///
/// assert_eq!(Test::with_text(CompareOp::Gt, "high", None).warning(), Some(Warning::NonNumeric));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Warning {
    /// A relational operator (`>`, `<`, `>=`, `<=`) with a literal that is
    /// neither a number nor a length.
    NonNumeric,
}

/// A unit suffix OfficeCLI's comparison recognizes (`ExtractUnit`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Unit {
    None,
    Pt,
    Px,
    Cm,
    In,
    Em,
    Percent,
}

/// What OfficeCLI's comparison reads from one side's text: its number with
/// one unit suffix stripped (`ExtractNumber`), that unit (`ExtractUnit`), and
/// its EMU value (`TryParseEmu`).
#[derive(Debug, Clone, Copy, PartialEq)]
struct Facts {
    number: Option<f64>,
    unit: Unit,
    emu: Option<i64>,
}

/// OfficeCLI's unit suffixes, in its matching order.
const UNIT_SUFFIXES: [(&str, Unit); 6] = [
    ("pt", Unit::Pt),
    ("px", Unit::Px),
    ("cm", Unit::Cm),
    ("in", Unit::In),
    ("em", Unit::Em),
    ("%", Unit::Percent),
];

impl Facts {
    fn of(text: &str) -> Self {
        Self {
            number: extract_number(text),
            unit: extract_unit(text),
            emu: try_emu(text),
        }
    }
}

/// `ExtractUnit`: the first recognized suffix of the untrimmed text.
fn extract_unit(text: &str) -> Unit {
    UNIT_SUFFIXES
        .iter()
        .find(|(suffix, _)| ends_with_ci(text, suffix))
        .map_or(Unit::None, |&(_, unit)| unit)
}

/// `ExtractNumber`: trim the end, strip the first recognized suffix, and read
/// a finite `NumericText` number.
fn extract_number(text: &str) -> Option<f64> {
    let trimmed = text.trim_end();
    let number = UNIT_SUFFIXES
        .iter()
        .find_map(|(suffix, _)| strip_suffix_ci(trimmed, suffix))
        .unwrap_or(trimmed);
    parse_numeric_text(number).filter(|value| value.is_finite())
}

/// `CompareNumeric`: numbers when the units agree or one side has none, else
/// EMU, else numbers; `None` when the sides are not comparable.
fn compare_numeric(actual: &Facts, expected: &Facts) -> Option<Ordering> {
    if let (Some(a), Some(b)) = (actual.number, expected.number) {
        if actual.unit == expected.unit || actual.unit == Unit::None || expected.unit == Unit::None
        {
            return a.partial_cmp(&b);
        }
    }
    if let (Some(a), Some(b)) = (actual.emu, expected.emu) {
        return Some(a.cmp(&b));
    }
    match (actual.number, expected.number) {
        (Some(a), Some(b)) => a.partial_cmp(&b),
        _ => None,
    }
}

/// `DimensionEquals`: EMU values within 500 when either side has a unit.
fn dimension_equals(actual: &Facts, expected: &Facts) -> bool {
    if actual.unit == Unit::None && expected.unit == Unit::None {
        return false;
    }
    match (actual.emu, expected.emu) {
        (Some(a), Some(b)) => (i128::from(a) - i128::from(b)).abs() <= DIMENSION_TOLERANCE,
        _ => false,
    }
}

/// `StringEquals`: case-insensitive, then again with leading `#`s removed
/// from both sides when either had one.
fn string_equals(actual: &str, expected: &str) -> bool {
    if eq_ignore_case(actual, expected) {
        return true;
    }
    let (a, b) = (
        actual.trim_start_matches('#'),
        expected.trim_start_matches('#'),
    );
    (a.len() != actual.len() || b.len() != expected.len()) && eq_ignore_case(a, b)
}

/// One query condition compiled once: its operator, the literal's text, what
/// OfficeCLI reads from the literal, and the literal coerced to the key's
/// type when it coerces.
///
/// | Operator | Holds when |
/// |---|---|
/// | `[key]` | the key is present and its text is not empty |
/// | `=` | the texts are equal ignoring case (and leading `#`); or both are lengths within 500 EMU and either names a unit; or both are equal numbers (a unit on one side only is ignored); or the literal coerces to the key's type and equals the value |
/// | `!=` | `=` does not hold; an absent key matches |
/// | `~=` | the value's text contains the literal, ignoring case; regex literals are refused |
/// | `>` `<` `>=` `<=` | numbers compare when the units agree or one side has none, else EMU values compare, else numbers; anything else never matches |
///
/// The value's text is its canonical form in the key's rendering, written
/// into a stack buffer, so evaluation never allocates.
///
/// ```
/// use litchi_core::selector::CompareOp;
/// use litchi_ops::value::{Color, ColorContext, LengthSpec, Length, Scalar, Test, ValueType};
///
/// let width = Some(ValueType::Length(LengthSpec::EXTENT));
/// let at_least = Test::with_text(CompareOp::Ge, "1in", width);
/// assert!(at_least.eval(Some(Scalar::Length(Length::from_emu(914_400)))));
/// assert!(!at_least.eval(None));
///
/// let fill = Some(ValueType::Color(ColorContext::Drawing));
/// let red = Test::with_text(CompareOp::Eq, "red", fill);
/// assert!(red.eval(Some(Scalar::Color(Color::rgb(255, 0, 0)))));
/// ```
#[derive(Clone, Copy)]
pub struct Test<'q> {
    compiled: Compiled,
    literal: &'q str,
}

impl<'q> Test<'q> {
    /// `[key]`: the key is present with non-empty text.
    #[must_use]
    pub const fn exists() -> Self {
        Self {
            compiled: Compiled::EXISTS,
            literal: "",
        }
    }

    /// Compiles `[key op literal]` for a key of type `ty` (`None` when the
    /// key's type is unknown). An operator this version does not know never
    /// matches; [`Test::compare`] refuses it instead.
    #[must_use]
    pub fn with_text(op: CompareOp, literal: &'q str, ty: Option<ValueType>) -> Self {
        Self {
            compiled: Compiled::new(op, literal, ty),
            literal,
        }
    }

    /// Compiles a grammar comparison.
    ///
    /// # Errors
    ///
    /// Returns `unsupported_capability` for a regular-expression literal
    /// (`~= r"..."`): the workspace has no regex engine (record 0010, open
    /// question Q3), and for an operator this version does not know.
    pub fn compare(op: CompareOp, value: &'q Value<'q>, ty: Option<ValueType>) -> Result<Self> {
        Ok(Self::with_text(op, literal(op, value)?, ty))
    }

    /// Compiles a grammar condition (`[key]` or `[key op value]`) for a key
    /// of type `ty`.
    ///
    /// # Errors
    ///
    /// As [`Test::compare`].
    pub fn from_condition(condition: &'q Condition<'q>, ty: Option<ValueType>) -> Result<Self> {
        match condition {
            Condition::Exists(_) => Ok(Self::exists()),
            Condition::Compare { op, value, .. } => Self::compare(*op, value, ty),
            _ => Err(Error::new(
                Code::UnsupportedCapability,
                "this kind of condition is not supported",
            )),
        }
    }

    /// The compile-time warning, if any.
    #[must_use]
    pub fn warning(&self) -> Option<Warning> {
        self.compiled.warning()
    }

    /// The literal read as a number the way OfficeCLI's header-row predicates
    /// decide "the literal is numeric" (`NumericText.TryParse`, no unit
    /// stripping).
    #[must_use]
    pub fn literal_number(&self) -> Option<f64> {
        parse_numeric_text(self.literal)
    }

    /// The literal's text.
    #[must_use]
    pub const fn literal(&self) -> &'q str {
        self.literal
    }

    /// Evaluates the condition against a key's value (`None` when the key is
    /// absent). Allocation-free.
    #[must_use]
    pub fn eval(&self, actual: Option<Scalar<'_>>) -> bool {
        self.compiled.eval(self.literal, actual)
    }
}

/// The literal of a grammar comparison, after refusing what query comparison
/// does not evaluate: a regular-expression literal (`~= r"..."`, record 0010
/// open question Q3) and an operator this version does not know.
pub(crate) fn literal<'v>(op: CompareOp, value: &'v Value<'_>) -> Result<&'v str> {
    if !matches!(
        op,
        CompareOp::Eq
            | CompareOp::Ne
            | CompareOp::Contains
            | CompareOp::Gt
            | CompareOp::Lt
            | CompareOp::Ge
            | CompareOp::Le
    ) {
        return Err(Error::new(
            Code::UnsupportedCapability,
            format!("the comparison operator '{op}' is not supported"),
        ));
    }
    match value {
        Value::Text(text) => Ok(text),
        Value::Regex(_) => Err(Error::new(
            Code::UnsupportedCapability,
            "regular-expression values are not supported",
        )
        .with_suggestion(
            "use '~=' with a plain substring; regex matching needs a regex engine the project has not adopted yet",
        )),
        _ => Err(Error::new(
            Code::UnsupportedCapability,
            "this kind of comparison value is not supported",
        )),
    }
}

/// A coerced literal that owns nothing it borrowed: every scalar but text.
/// Query literals never coerce to [`Scalar::Text`] (free text compares as
/// text), so nothing is lost.
fn detach(scalar: Scalar<'_>) -> Option<Scalar<'static>> {
    Some(match scalar {
        Scalar::Bool(value) => Scalar::Bool(value),
        Scalar::Int(value) => Scalar::Int(value),
        Scalar::Number(value) => Scalar::Number(value),
        Scalar::Token(token) => Scalar::Token(token),
        Scalar::Length(length) => Scalar::Length(length),
        Scalar::Color(color) => Scalar::Color(color),
        Scalar::Text(_) => return None,
    })
}

/// A [`Test`] without its literal text: the operator, the key's type, what
/// OfficeCLI reads from the literal, and the literal coerced to the key's
/// type. A compiled query plan that owns its literals keeps this and passes
/// the literal back at evaluation, so the plan needs no self-reference.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Compiled {
    op: Op,
    ty: Option<ValueType>,
    facts: Facts,
    typed: Option<Scalar<'static>>,
}

impl Compiled {
    /// `[key]`: present with non-empty text.
    pub(crate) const EXISTS: Self = Self {
        op: Op::Exists,
        ty: None,
        facts: Facts {
            number: None,
            unit: Unit::None,
            emu: None,
        },
        typed: None,
    };

    /// Compiles `[key op literal]` for a key of type `ty`.
    pub(crate) fn new(op: CompareOp, literal: &str, ty: Option<ValueType>) -> Self {
        let op = match op {
            CompareOp::Eq => Op::Eq,
            CompareOp::Ne => Op::Ne,
            CompareOp::Contains => Op::Contains,
            CompareOp::Gt => Op::Gt,
            CompareOp::Lt => Op::Lt,
            CompareOp::Ge => Op::Ge,
            CompareOp::Le => Op::Le,
            // `literal` refuses an operator this version does not know;
            // compiled here, it never matches.
            _ => Op::Never,
        };
        let typed = match (op, ty) {
            (Op::Eq | Op::Ne, Some(ty)) if ty != ValueType::Text && !literal.trim().is_empty() => {
                ty.coerce(literal).ok().and_then(detach)
            },
            _ => None,
        };
        Self {
            op,
            ty,
            facts: Facts::of(literal),
            typed,
        }
    }

    /// `Some(true)` for `=`, `Some(false)` for `!=`, `None` otherwise.
    pub(crate) const fn polarity(&self) -> Option<bool> {
        match self.op {
            Op::Eq => Some(true),
            Op::Ne => Some(false),
            _ => None,
        }
    }

    /// Whether the operator is relational (`>`, `<`, `>=`, `<=`).
    pub(crate) const fn is_relational(&self) -> bool {
        matches!(self.op, Op::Gt | Op::Lt | Op::Ge | Op::Le)
    }

    /// The compile-time warning, if any.
    pub(crate) fn warning(&self) -> Option<Warning> {
        (self.is_relational() && self.facts.number.is_none() && self.facts.emu.is_none())
            .then_some(Warning::NonNumeric)
    }

    /// Evaluates the condition, whose literal is `literal`, against a key's
    /// value (`None` when the key is absent). Allocation-free.
    pub(crate) fn eval(&self, literal: &str, actual: Option<Scalar<'_>>) -> bool {
        let Some(actual) = actual else {
            return self.op == Op::Ne;
        };
        let mut buffer = StackStr::<DISPLAY_CAPACITY>::new();
        let text = match actual {
            Scalar::Text(text) => text,
            _ => {
                if write!(buffer, "{}", actual.display(self.ty)).is_err() {
                    // Unreachable for the canonical forms, which all fit.
                    return self.op == Op::Ne;
                }
                buffer.as_str()
            },
        };
        match self.op {
            Op::Exists => !text.is_empty(),
            Op::Eq => self.equals(literal, &actual, text),
            Op::Ne => !self.equals(literal, &actual, text),
            Op::Contains => contains_ignore_case(text, literal),
            Op::Gt => self.relation(text, Ordering::is_gt),
            Op::Lt => self.relation(text, Ordering::is_lt),
            Op::Ge => self.relation(text, Ordering::is_ge),
            Op::Le => self.relation(text, Ordering::is_le),
            Op::Never => false,
        }
    }

    /// Whether the literal reads as neither a number nor a length, so no
    /// numeric or dimension comparison can hold whatever the value is.
    fn literal_is_text(&self) -> bool {
        self.facts.number.is_none() && self.facts.emu.is_none()
    }

    fn relation(&self, text: &str, holds: fn(Ordering) -> bool) -> bool {
        !self.literal_is_text() && compare_numeric(&Facts::of(text), &self.facts).is_some_and(holds)
    }

    fn equals(&self, literal: &str, actual: &Scalar<'_>, text: &str) -> bool {
        if string_equals(text, literal) {
            return true;
        }
        if self.literal_is_text() {
            // Reading the value's number and EMU could not change the
            // answer: both comparisons need them on both sides.
            return self.typed_equals(actual);
        }
        let facts = Facts::of(text);
        dimension_equals(&facts, &self.facts)
            || compare_numeric(&facts, &self.facts) == Some(Ordering::Equal)
            || self.typed_equals(actual)
    }

    /// Typed equality: the literal coerced to the key's type equals the value.
    fn typed_equals(&self, actual: &Scalar<'_>) -> bool {
        match (self.typed, *actual) {
            (Some(Scalar::Bool(a)), Scalar::Bool(b)) => a == b,
            (Some(Scalar::Int(a)), Scalar::Int(b)) => a == b,
            (Some(Scalar::Number(a)), Scalar::Number(b)) => a == b,
            (Some(Scalar::Token(a)), Scalar::Token(b) | Scalar::Text(b)) => {
                a.eq_ignore_ascii_case(b.trim())
            },
            (Some(Scalar::Length(a)), Scalar::Length(b)) => a == b,
            (Some(Scalar::Color(a)), Scalar::Color(b)) => a.matches(&b),
            _ => false,
        }
    }
}

impl std::fmt::Debug for Test<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Test").finish_non_exhaustive()
    }
}
