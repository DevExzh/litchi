//! Query comparison: one `[key op value]` condition compiled once and
//! evaluated against typed values without allocating.
//!
//! The semantics are OfficeCLI's `AttributeFilter.MatchOne`
//! (`OC/Core/AttributeFilter.cs`), applied to the canonical text of the
//! node's typed value, plus typed equality: a literal that coerces to the
//! key's type equals a value it coerces to (`[fill=red]` matches `#FF0000`).
//!
//! Evaluation reads only what the operator needs (record 0262). What
//! OfficeCLI reads from the canonical text (its number, unit and EMU value)
//! is taken from the typed value directly where the text is known to read
//! back to it: an integer, a finite number, a length in its key's style.
//! Text is written only for a substring test, for colours, and for an
//! equality whose literal could match the text without matching its number
//! (a leading `#`, or non-ASCII letters that fold to ASCII). A text's EMU
//! value is read only when a comparison gets that far. The result is always
//! the one the written text gives; the tests check that against the
//! write-then-read evaluation.

use std::cmp::Ordering;
use std::fmt::Write as _;

use litchi_core::selector::{CompareOp, Condition, Value};

use crate::error::{Code, Error, Result};

use super::length::{DisplayUnit, try_emu};
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
    /// What OfficeCLI reads from `text`.
    fn of(text: &str) -> Self {
        let (number, unit) = number_and_unit(text);
        Self {
            number,
            unit,
            emu: text_emu(text),
        }
    }

    /// Nothing: the facts of text without a digit, such as `true`.
    const NONE: Self = Self {
        number: None,
        unit: Unit::None,
        emu: None,
    };
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

/// `ExtractNumber` and `ExtractUnit` of `text`, with a shortcut that gives
/// the same answer: a plain decimal (`-?[0-9]*.?[0-9]*` with a digit) has no
/// unit and is the number a correctly rounded parse reads, as `NumericText`
/// reads it. (Text without a digit is [`Lazy::text`]'s shortcut.)
fn number_and_unit(text: &str) -> (Option<f64>, Unit) {
    match shape(text) {
        Shape::NoDigit => (None, extract_unit(text)),
        Shape::Decimal => (decimal(text).filter(|value| value.is_finite()), Unit::None),
        Shape::Other => (extract_number(text), extract_unit(text)),
    }
}

/// The correctly rounded double of a plain decimal (`-?[0-9]*.?[0-9]*` with
/// a digit). Up to 2^53 in its digits and 22 decimals, one division of exact
/// doubles rounds correctly (Clinger's fast path); longer decimals go
/// through Rust's parse, which rounds correctly too.
fn decimal(text: &str) -> Option<f64> {
    const POWERS: [f64; 23] = [
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15, 1e16,
        1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
    ];
    let bytes = text.as_bytes();
    let (negative, digits) = match bytes.split_first() {
        Some((b'-', rest)) => (true, rest),
        _ => (false, bytes),
    };
    let (mut mantissa, mut decimals, mut point) = (0_u64, 0_usize, false);
    for &byte in digits {
        if byte == b'.' {
            point = true;
            continue;
        }
        let digit = u64::from(byte.wrapping_sub(b'0'));
        match mantissa
            .checked_mul(10)
            .and_then(|value| value.checked_add(digit))
        {
            Some(value) if value <= 1 << 53 => mantissa = value,
            _ => return text.parse().ok(),
        }
        decimals += usize::from(point);
    }
    let power = POWERS.get(decimals).copied();
    let Some(power) = power else {
        return text.parse().ok();
    };
    #[allow(
        clippy::cast_precision_loss,
        reason = "the mantissa is at most 2^53, exact as a double"
    )]
    let value = mantissa as f64 / power;
    Some(if negative { -value } else { value })
}

/// `TryParseEmu` of `text`: every route of `ParseEmu` needs a digit.
fn text_emu(text: &str) -> Option<i64> {
    if text.bytes().any(|byte| byte.is_ascii_digit()) {
        try_emu(text)
    } else {
        None
    }
}

/// How `number_and_unit` reads a text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Shape {
    /// No ASCII digit: neither a number nor an EMU value.
    NoDigit,
    /// `-?[0-9]*.?[0-9]*` with at least one digit.
    Decimal,
    /// Anything else.
    Other,
}

fn shape(text: &str) -> Shape {
    let bytes = text.as_bytes();
    let body = bytes.strip_prefix(b"-").unwrap_or(bytes);
    let (mut digits, mut points, mut other) = (false, 0_usize, false);
    for &byte in body {
        match byte {
            b'0'..=b'9' => digits = true,
            b'.' => points += 1,
            _ => other = true,
        }
    }
    if digits && !other && points <= 1 {
        Shape::Decimal
    } else if digits || bytes.iter().any(u8::is_ascii_digit) {
        Shape::Other
    } else {
        Shape::NoDigit
    }
}

/// One side's facts, read as a comparison needs them: the number and unit
/// first, the EMU value only when a comparison reaches it.
struct Lazy<'t> {
    head: Option<(Option<f64>, Unit)>,
    emu: Option<Option<i64>>,
    /// Where the facts come from.
    source: Source<'t>,
}

/// Where a [`Lazy`] reads its facts.
enum Source<'t> {
    /// Facts already known.
    Known,
    /// Read from this text.
    Text(&'t str),
    /// A finite number: its text is .NET's round-trip rendering, which reads
    /// back as the number itself with no unit; its EMU value is read from
    /// that text when needed.
    Number(f64),
}

impl<'t> Lazy<'t> {
    const fn known(facts: Facts) -> Self {
        Self {
            head: Some((facts.number, facts.unit)),
            emu: Some(facts.emu),
            source: Source::Known,
        }
    }

    fn text(text: &'t str) -> Self {
        if !text.bytes().any(|byte| byte.is_ascii_digit()) {
            // Without a digit there is neither a number nor an EMU value,
            // and no comparison reads the unit of a side that has neither.
            return Self::known(Facts::NONE);
        }
        Self {
            head: None,
            emu: None,
            source: Source::Text(text),
        }
    }

    const fn number(value: f64) -> Self {
        Self {
            head: Some((Some(value), Unit::None)),
            emu: None,
            source: Source::Number(value),
        }
    }

    /// The number and the unit.
    fn head(&mut self) -> (Option<f64>, Unit) {
        if let Some(head) = self.head {
            return head;
        }
        let head = match self.source {
            Source::Text(text) => number_and_unit(text),
            Source::Known | Source::Number(_) => (None, Unit::None),
        };
        self.head = Some(head);
        head
    }

    /// The EMU value.
    fn emu(&mut self) -> Option<i64> {
        if let Some(emu) = self.emu {
            return emu;
        }
        let emu = match self.source {
            Source::Text(text) => text_emu(text),
            Source::Number(value) => {
                let mut buffer = StackStr::<DISPLAY_CAPACITY>::new();
                match write!(buffer, "{}", Scalar::Number(value).display(None)) {
                    Ok(()) => text_emu(buffer.as_str()),
                    Err(_) => None,
                }
            },
            Source::Known => None,
        };
        self.emu = Some(emu);
        emu
    }
}

/// `CompareNumeric`: numbers when the units agree or one side has none, else
/// EMU, else numbers; `None` when the sides are not comparable.
fn compare_numeric(actual: &mut Lazy<'_>, expected: &Facts) -> Option<Ordering> {
    let (number, unit) = actual.head();
    if let (Some(a), Some(b)) = (number, expected.number) {
        if unit == expected.unit || unit == Unit::None || expected.unit == Unit::None {
            return a.partial_cmp(&b);
        }
    }
    if let Some(b) = expected.emu {
        if let Some(a) = actual.emu() {
            return Some(a.cmp(&b));
        }
    }
    match (number, expected.number) {
        (Some(a), Some(b)) => a.partial_cmp(&b),
        _ => None,
    }
}

/// `DimensionEquals`: EMU values within 500 when either side has a unit.
fn dimension_equals(actual: &mut Lazy<'_>, expected: &Facts) -> bool {
    let (_, unit) = actual.head();
    if unit == Unit::None && expected.unit == Unit::None {
        return false;
    }
    match expected.emu {
        Some(b) => actual
            .emu()
            .is_some_and(|a| (i128::from(a) - i128::from(b)).abs() <= DIMENSION_TOLERANCE),
        None => false,
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

/// The canonical text of a value that has one without writing it.
fn borrowed_text<'a>(actual: &Scalar<'a>) -> Option<&'a str> {
    match *actual {
        Scalar::Text(text) => Some(text),
        Scalar::Token(token) => Some(token),
        Scalar::Bool(value) => Some(if value { "true" } else { "false" }),
        _ => None,
    }
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
        Scalar::Text(_) | Scalar::List(_) => return None,
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
    /// The literal is ASCII and does not start with `#`: it equals the
    /// canonical text of an integer, a number or a length only when its
    /// number or EMU value equals the value's, so equality need not write
    /// that text.
    plain: bool,
}

impl Compiled {
    /// `[key]`: present with non-empty text.
    pub(crate) const EXISTS: Self = Self {
        op: Op::Exists,
        ty: None,
        facts: Facts::NONE,
        typed: None,
        plain: true,
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
            plain: literal.is_ascii() && !literal.starts_with('#'),
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
    ///
    /// The result is OfficeCLI's on the value's canonical text (see the
    /// module documentation for what is read without writing it).
    pub(crate) fn eval(&self, literal: &str, actual: Option<Scalar<'_>>) -> bool {
        let Some(actual) = actual else {
            return self.op == Op::Ne;
        };
        // A list (the keyword list, record 0354) holds when one of its
        // items does; `!=` when none equals; `exists` when it has an item.
        if let Scalar::List(items) = actual {
            return match self.op {
                Op::Exists => !items.is_empty(),
                Op::Ne => !items
                    .iter()
                    .any(|value| self.equals(literal, &Scalar::Text(value))),
                _ => items
                    .iter()
                    .any(|value| self.eval(literal, Some(Scalar::Text(value)))),
            };
        }
        match self.op {
            Op::Eq => self.equals(literal, &actual),
            Op::Ne => !self.equals(literal, &actual),
            Op::Gt => self.relation(&actual, Ordering::is_gt),
            Op::Lt => self.relation(&actual, Ordering::is_lt),
            Op::Ge => self.relation(&actual, Ordering::is_ge),
            Op::Le => self.relation(&actual, Ordering::is_le),
            Op::Exists => match actual {
                // The canonical text of a boolean, an integer, a finite number
                // or a length is never empty.
                Scalar::Bool(_) | Scalar::Int(_) | Scalar::Length(_) => true,
                Scalar::Number(value) if value.is_finite() => true,
                _ => self
                    .with_text(&actual, |text| !text.is_empty())
                    .unwrap_or(false),
            },
            Op::Contains => self
                .with_text(&actual, |text| contains_ignore_case(text, literal))
                .unwrap_or(false),
            Op::Never => false,
        }
    }

    /// Runs `read` on the value's canonical text, written into a stack buffer
    /// when the value does not hold it; `None` if it does not fit, which no
    /// canonical form reaches (the longest, a colour with eight transforms,
    /// takes about 160 bytes).
    fn with_text<T>(&self, actual: &Scalar<'_>, read: impl FnOnce(&str) -> T) -> Option<T> {
        if let Some(text) = borrowed_text(actual) {
            return Some(read(text));
        }
        let mut buffer = StackStr::<DISPLAY_CAPACITY>::new();
        write!(buffer, "{}", actual.display(self.ty)).ok()?;
        Some(read(buffer.as_str()))
    }

    /// What OfficeCLI reads from the canonical text of an integer, a finite
    /// number or a length, known without writing the text; `None` for the
    /// other values.
    fn typed_facts(&self, actual: &Scalar<'_>) -> Option<Lazy<'static>> {
        Some(match *actual {
            // `i64` text reads back exactly: `NumericText` gives the
            // correctly rounded double, `ParseEmu` the integer.
            Scalar::Int(value) => Lazy::known(Facts {
                #[allow(
                    clippy::cast_precision_loss,
                    reason = "the correctly rounded double is what the text reads as"
                )]
                number: Some(value as f64),
                unit: Unit::None,
                emu: Some(value),
            }),
            // .NET's round-trip rendering reads back as the number itself.
            Scalar::Number(value) if value.is_finite() => Lazy::number(value),
            Scalar::Length(length) => {
                let facts = length.display_facts(ValueType::length_style(self.ty))?;
                Lazy::known(Facts {
                    number: facts.number,
                    unit: match facts.unit {
                        Some(DisplayUnit::Cm) => Unit::Cm,
                        Some(DisplayUnit::Pt) => Unit::Pt,
                        Some(DisplayUnit::In) => Unit::In,
                        None => Unit::None,
                    },
                    emu: facts.emu,
                })
            },
            _ => return None,
        })
    }

    /// Whether the literal reads as neither a number nor a length, so no
    /// numeric or dimension comparison can hold whatever the value is.
    fn literal_is_text(&self) -> bool {
        self.facts.number.is_none() && self.facts.emu.is_none()
    }

    fn relation(&self, actual: &Scalar<'_>, holds: fn(Ordering) -> bool) -> bool {
        if self.literal_is_text() {
            return false;
        }
        if let Some(mut facts) = self.typed_facts(actual) {
            return compare_numeric(&mut facts, &self.facts).is_some_and(holds);
        }
        if let Some(text) = borrowed_text(actual) {
            return compare_numeric(&mut Lazy::text(text), &self.facts).is_some_and(holds);
        }
        self.with_text(actual, |text| {
            compare_numeric(&mut Lazy::text(text), &self.facts).is_some_and(holds)
        })
        .unwrap_or(false)
    }

    fn equals(&self, literal: &str, actual: &Scalar<'_>) -> bool {
        if self.plain {
            if let Some(mut facts) = self.typed_facts(actual) {
                // A plain literal equal to the value's text, ignoring ASCII
                // case, reads as the same number and EMU value, which the
                // canonical text of these values always has: the numeric
                // comparison below finds it, so the text is not written.
                return self.equals_facts(&mut facts, actual);
            }
        }
        if let Some(text) = borrowed_text(actual) {
            return self.equals_text(literal, actual, text);
        }
        self.with_text(actual, |text| self.equals_text(literal, actual, text))
            .unwrap_or(false)
    }

    /// `=` on the value's canonical text.
    fn equals_text(&self, literal: &str, actual: &Scalar<'_>, text: &str) -> bool {
        string_equals(text, literal) || self.equals_facts(&mut Lazy::text(text), actual)
    }

    /// `=` past the text comparison: dimensions, numbers, then the typed
    /// value.
    fn equals_facts(&self, facts: &mut Lazy<'_>, actual: &Scalar<'_>) -> bool {
        if self.literal_is_text() {
            // Reading the value's number and EMU could not change the
            // answer: both comparisons need them on both sides.
            return self.typed_equals(actual);
        }
        dimension_equals(facts, &self.facts)
            || compare_numeric(facts, &self.facts) == Some(Ordering::Equal)
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

/// The evaluation before record 0262, kept verbatim as the oracle the tests
/// compare [`Compiled::eval`] with: write the value's canonical text, then
/// read everything OfficeCLI reads from it.
#[cfg(test)]
pub(crate) mod reference {
    use std::cmp::Ordering;
    use std::fmt::Write as _;

    use super::{Compiled, DIMENSION_TOLERANCE, DISPLAY_CAPACITY, Facts, Op, Unit, string_equals};
    use crate::value::text::{StackStr, contains_ignore_case};
    use crate::value::ty::Scalar;

    fn compare_numeric(actual: &Facts, expected: &Facts) -> Option<Ordering> {
        if let (Some(a), Some(b)) = (actual.number, expected.number) {
            if actual.unit == expected.unit
                || actual.unit == Unit::None
                || expected.unit == Unit::None
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

    fn dimension_equals(actual: &Facts, expected: &Facts) -> bool {
        if actual.unit == Unit::None && expected.unit == Unit::None {
            return false;
        }
        match (actual.emu, expected.emu) {
            (Some(a), Some(b)) => (i128::from(a) - i128::from(b)).abs() <= DIMENSION_TOLERANCE,
            _ => false,
        }
    }

    impl Compiled {
        /// [`Compiled::eval`] as it was: write, then read.
        pub(crate) fn reference_eval(&self, literal: &str, actual: Option<Scalar<'_>>) -> bool {
            let Some(actual) = actual else {
                return self.op == Op::Ne;
            };
            let mut buffer = StackStr::<DISPLAY_CAPACITY>::new();
            let text = match actual {
                Scalar::Text(text) => text,
                _ => {
                    if write!(buffer, "{}", actual.display(self.ty)).is_err() {
                        return self.op == Op::Ne;
                    }
                    buffer.as_str()
                },
            };
            let literal_is_text = self.facts.number.is_none() && self.facts.emu.is_none();
            let relation = |holds: fn(Ordering) -> bool| {
                !literal_is_text
                    && compare_numeric(&Facts::of_reference(text), &self.facts).is_some_and(holds)
            };
            let equals = || {
                if string_equals(text, literal) {
                    return true;
                }
                if literal_is_text {
                    return self.typed_equals(&actual);
                }
                let facts = Facts::of_reference(text);
                dimension_equals(&facts, &self.facts)
                    || compare_numeric(&facts, &self.facts) == Some(Ordering::Equal)
                    || self.typed_equals(&actual)
            };
            match self.op {
                Op::Exists => !text.is_empty(),
                Op::Eq => equals(),
                Op::Ne => !equals(),
                Op::Contains => contains_ignore_case(text, literal),
                Op::Gt => relation(Ordering::is_gt),
                Op::Lt => relation(Ordering::is_lt),
                Op::Ge => relation(Ordering::is_ge),
                Op::Le => relation(Ordering::is_le),
                Op::Never => false,
            }
        }
    }

    impl Facts {
        /// `Facts::of` as it was: no shortcuts.
        pub(crate) fn of_reference(text: &str) -> Self {
            Self {
                number: super::extract_number(text),
                unit: super::extract_unit(text),
                emu: crate::value::length::try_emu(text),
            }
        }
    }
}
