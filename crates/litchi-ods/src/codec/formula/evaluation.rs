//! Bounded, read-only scalar evaluation for OpenFormula expressions.
//!
//! This module evaluates the already-parsed ODS expression tree directly.  It
//! does not stringify and reparse formulas, inspect a workbook, resolve an
//! external source, refresh a cache, or publish a cell change.  The scalar
//! profile covers constants, scalar operators, the normative logical
//! functions (`TRUE`, `FALSE`, `IF`, `IFERROR`, `IFNA`, `AND`, `OR`, `NOT`,
//! and `XOR`), the five bit-operation functions (`BITAND`, `BITLSHIFT`,
//! `BITOR`, `BITRSHIFT`, and `BITXOR`), and the section 6.19 radix functions
//! (`BASE`, `DECIMAL`, and the twelve `xxx2yyy` conversions).  References,
//! arrays, names, labels, and other functions are reported as typed
//! capability refusals.
//!
//! The profile makes the following deterministic choices for host-dependent
//! scalar behavior:
//!
//! * all numbers are finite `f64` values; non-finite literals and results are
//!   returned as a Number error;
//! * number equality is exact `f64` equality;
//! * text is compared without Unicode normalization, using case-sensitive
//!   Unicode scalar ordering by default;
//! * this bounded profile uses explicit case-sensitive text comparison; a
//!   future host setting must use a specified Unicode case-folding policy
//!   before being exposed here;
//! * text-to-number conversion accepts only locale-independent decimal-point
//!   syntax and reports a value error when conversion fails;
//! * Text used where a Logical is required reports `#VALUE!`; AND/OR follow
//!   their `NumberSequenceList` signature and convert scalar numeric Text;
//! * bit-operation data operands and results use the exact unsigned 48-bit
//!   range; signed finite shift counts have no arbitrary magnitude cap.  A
//!   left shift that cannot fit returns `#NUM!`, while a sufficiently large
//!   right shift returns zero.  Integer conversion truncates toward zero;
//!   unrepresentable data operands/results return `#NUM!` in this profile;
//! * `0^0` is accepted as `1`, as permitted by this bounded profile.
//! * `BASE` and `DECIMAL` use a fixed 1024-bit unsigned magnitude so every
//!   finite integer `f64` can be converted without narrowing through `u64`.
//!   Decimal text is accumulated once and rounded to `f64` using a
//!   nearest-even conversion.  The direct binary, octal, and hexadecimal
//!   conversions use the specified 10-, 30-, and 40-bit two's-complement
//!   widths and uppercase text output.  Their optional `Digits` argument is
//!   truncated toward zero, accepts values from zero through ten, pads
//!   positive results, and is ignored for negative results after any formula
//!   error in the argument has been propagated.
//!
//! A successful text result keeps its memory reservation in
//! [`EvaluatedScalar`](crate::codec::formula::evaluation::EvaluatedScalar) until that result is dropped. Borrowed source strings
//! do not allocate or require an output reservation.
//! Evaluation limits charge AST visits, operator applications, and admitted
//! text/numeric byte work. The storage limit covers aggregate live requested
//! heap capacity for the evaluator stacks and owned text, including capacity
//! growth. Source-tree storage and caller-created copies are outside that
//! limit. Stack/vector/string growth uses fallible reservations; the shared
//! core budget and error bookkeeping retain their existing allocation behavior.
//!
//! With a caller-supplied execution context:
//!
//! ```rust
//! use litchi_ods::codec::formula::{
//!     expression::Expression,
//!     evaluation::{evaluate_scalar_with_context, EvaluationLimits, ScalarValue},
//! };
//! # use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Profile};
//! # use std::num::{NonZeroU64, NonZeroUsize};
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! # let (_cancellation, token) = CancellationSource::pair();
//! # let execution = ExecutionContext::new(
//! #     Budget::root("formula", Limits::for_profile(Profile::Server)), token,
//! #     ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)?,
//! # );
//! // The reference in the unselected branch is never resolved, and IFERROR
//! // catches the selected branch's formula-level division error.
//! let expression = Expression::parse("of:=IF(FALSE();[.A1];IFERROR(1/0;42))")?;
//! let result = evaluate_scalar_with_context(
//!     &expression, &execution, &EvaluationLimits::default(),
//! )?;
//! assert_eq!(result.value(), &ScalarValue::Number(42.0));
//! # Ok(())
//! # }
//! ```

mod radix;

use super::expression::{Expression, InfixOperator, Kind, Node, PostfixOperator, PrefixOperator};
use litchi_core::{Budget, ExecutionContext, ExecutionError, Reservation, Resource, ResourceLimit};
use std::{
    borrow::Cow,
    cmp::Ordering,
    error::Error as StdError,
    fmt::{self, Display, Formatter, Write as FmtWrite},
    mem::size_of,
    sync::Arc,
};

/// Default maximum number of evaluator steps for one scalar operation.
pub const DEFAULT_MAX_EVALUATION_STEPS: u64 = 1_000_000;
/// Default maximum UTF-8 bytes in one scalar text value.
pub const DEFAULT_MAX_TEXT_BYTES: usize = 32_767;
/// Default maximum bytes reserved for evaluator-owned temporary/result storage.
pub const DEFAULT_MAX_STORAGE_BYTES: usize = 4 * 1024 * 1024;
/// Default maximum entries in either explicit evaluator stack.
pub const DEFAULT_MAX_STACK_ENTRIES: usize = 65_536;

const EVALUATION_SCOPE: &str = "ods-formula-evaluation";
// Rust's general `Display` formatter may choose fixed notation for very large
// or very small finite values. Keep formatting fallible without a temporary
// heap allocation; this covers the finite `f64` display range with room for
// sign and punctuation.
const MAX_NUMBER_TEXT_BYTES: usize = 1024;
const STRING_COPY_CHUNK_BYTES: usize = 4096;

/// Finite limits for one scalar evaluation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EvaluationLimits {
    max_steps: u64,
    max_text_bytes: usize,
    max_storage_bytes: usize,
    max_stack_entries: usize,
}

impl Default for EvaluationLimits {
    fn default() -> Self {
        Self {
            max_steps: DEFAULT_MAX_EVALUATION_STEPS,
            max_text_bytes: DEFAULT_MAX_TEXT_BYTES,
            max_storage_bytes: DEFAULT_MAX_STORAGE_BYTES,
            max_stack_entries: DEFAULT_MAX_STACK_ENTRIES,
        }
    }
}

impl EvaluationLimits {
    /// Set the maximum combined AST, operator, and byte-work units.
    #[must_use]
    pub const fn with_max_steps(mut self, value: u64) -> Self {
        self.max_steps = value;
        self
    }

    /// Set the maximum UTF-8 byte length of a scalar text value.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, value: usize) -> Self {
        self.max_text_bytes = value;
        self
    }

    /// Set the maximum bytes reserved by evaluator-owned storage.
    #[must_use]
    pub const fn with_max_storage_bytes(mut self, value: usize) -> Self {
        self.max_storage_bytes = value;
        self
    }

    /// Set the maximum number of entries in either explicit evaluator stack.
    #[must_use]
    pub const fn with_max_stack_entries(mut self, value: usize) -> Self {
        self.max_stack_entries = value;
        self
    }

    /// Return the maximum number of evaluation steps.
    #[must_use]
    pub const fn max_steps(self) -> u64 {
        self.max_steps
    }

    /// Return the maximum text byte length.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Return the maximum evaluator-owned storage bytes.
    #[must_use]
    pub const fn max_storage_bytes(self) -> usize {
        self.max_storage_bytes
    }

    /// Return the maximum explicit stack size.
    #[must_use]
    pub const fn max_stack_entries(self) -> usize {
        self.max_stack_entries
    }
}

/// Text comparison policy used by the scalar profile.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextCase {
    /// Compare source Unicode scalar values exactly.
    Sensitive,
}

/// Locale-independent scalar evaluation options.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EvaluationOptions {
    text_case: TextCase,
}

impl Default for EvaluationOptions {
    fn default() -> Self {
        Self {
            text_case: TextCase::Sensitive,
        }
    }
}

impl EvaluationOptions {
    /// Return the configured text comparison policy.
    #[must_use]
    pub const fn text_case(self) -> TextCase {
        self.text_case
    }
}

/// Evaluation context carrying the caller-owned execution policy.
#[derive(Clone, Copy, Debug)]
pub struct EvaluationContext<'a> {
    execution: &'a ExecutionContext,
    options: EvaluationOptions,
}

impl<'a> EvaluationContext<'a> {
    /// Create a context using deterministic default options.
    #[must_use]
    pub const fn new(execution: &'a ExecutionContext) -> Self {
        Self {
            execution,
            options: EvaluationOptions {
                text_case: TextCase::Sensitive,
            },
        }
    }

    /// Create a context with explicit scalar options.
    #[must_use]
    pub const fn with_options(execution: &'a ExecutionContext, options: EvaluationOptions) -> Self {
        Self { execution, options }
    }

    /// Return the caller-owned execution context.
    #[must_use]
    pub const fn execution(self) -> &'a ExecutionContext {
        self.execution
    }

    /// Return the scalar options.
    #[must_use]
    pub const fn options(self) -> EvaluationOptions {
        self.options
    }
}

/// A formula-level error value.  Errors are values and do not become Rust
/// failures during ordinary operator propagation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ScalarError {
    /// The OpenFormula `#N/A` value.
    NotAvailable,
    /// An unknown or invalid error spelling, represented as `#NAME?`.
    Name,
    /// A type, domain, or mixed-comparison error.
    Value,
    /// Division by zero.
    DivisionByZero,
    /// A reference-related error value.
    Reference,
    /// A numeric domain or non-finite-result error.
    Number,
    /// The OpenFormula `#NULL!` value.
    Null,
}

impl Display for ScalarError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::NotAvailable => "#N/A",
            Self::Name => "#NAME?",
            Self::Value => "#VALUE!",
            Self::DivisionByZero => "#DIV/0!",
            Self::Reference => "#REF!",
            Self::Number => "#NUM!",
            Self::Null => "#NULL!",
        })
    }
}

/// Scalar value returned by the ODS evaluation profile.
#[derive(Debug, PartialEq)]
#[non_exhaustive]
pub enum ScalarValue<'a> {
    /// A finite OpenFormula Number.
    Number(f64),
    /// A logical value distinct from Number in this profile.
    Logical(bool),
    /// Text borrowed from the immutable expression where possible.
    Text(Cow<'a, str>),
    /// A formula-level Error value.
    Error(ScalarError),
}

/// A successful scalar result.  Owned text retains its memory reservation
/// until this value is dropped.  Use [`Self::value`] for a borrowed view.
#[derive(Debug)]
pub struct EvaluatedScalar<'a> {
    value: ScalarValue<'a>,
    output_reservation: Option<Reservation>,
}

impl<'a> EvaluatedScalar<'a> {
    fn new(value: ScalarValue<'a>, output_reservation: Option<Reservation>) -> Self {
        Self {
            value,
            output_reservation,
        }
    }

    /// Borrow the evaluated scalar value.
    #[must_use]
    pub const fn value(&self) -> &ScalarValue<'a> {
        &self.value
    }

    /// Return the owned output reservation amount, if any.
    #[must_use]
    pub fn reserved_output_bytes(&self) -> usize {
        self.output_reservation
            .as_ref()
            .and_then(|reservation| usize::try_from(reservation.amount()).ok())
            .unwrap_or(0)
    }
}

/// Capability refusals from the bounded scalar profile.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum UnsupportedKind {
    /// A bracketed or rich reference requires a resolver/context.
    Reference,
    /// A range, intersection, or union reference operator requires a resolver.
    ReferenceOperator,
    /// An inline array requires non-scalar evaluation.
    Array,
    /// A named expression requires a definition resolver.
    NamedExpression,
    /// A quoted label or automatic intersection requires a label resolver.
    Label,
    /// A missing function parameter outside a function's defined defaults.
    MissingArgument,
    /// A standard or host-defined function outside this scalar profile.
    Function,
}

impl Display for UnsupportedKind {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Reference => "ODS scalar evaluation does not resolve references",
            Self::ReferenceOperator => "ODS scalar evaluation does not resolve reference operators",
            Self::Array => "ODS scalar evaluation does not evaluate arrays",
            Self::NamedExpression => "ODS scalar evaluation does not resolve named expressions",
            Self::Label => "ODS scalar evaluation does not resolve labels",
            Self::MissingArgument => "ODS scalar evaluation does not evaluate missing arguments",
            Self::Function => "ODS scalar evaluation does not implement this function",
        })
    }
}

/// Failures of the evaluator operation itself.  Formula-level errors are
/// represented by [`ScalarValue::Error`] instead.
#[derive(Debug)]
#[non_exhaustive]
pub enum EvaluationFailure {
    /// A value requires a capability outside this profile.
    Unsupported(UnsupportedKind),
    /// The immutable AST or evaluator invariant was inconsistent.
    InvalidExpression(&'static str),
    /// The caller's cancellation token was set.
    Cancelled,
    /// A finite local or hierarchical resource limit was exceeded.
    ResourceLimit(ResourceLimit),
    /// A fallible evaluator-owned allocation could not be admitted.
    Allocation {
        /// Resource description for the failed allocation.
        resource: &'static str,
        /// Original allocator error.
        source: std::collections::TryReserveError,
    },
    /// A non-limit execution policy error from the caller context.
    Execution(ExecutionError),
}

impl Display for EvaluationFailure {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> fmt::Result {
        match self {
            Self::Unsupported(kind) => kind.fmt(formatter),
            Self::InvalidExpression(message) => formatter.write_str(message),
            Self::Cancelled => formatter.write_str("ODS scalar evaluation cancelled"),
            Self::ResourceLimit(limit) => limit.fmt(formatter),
            Self::Allocation { resource, source } => {
                write!(formatter, "allocation failed for {resource}: {source}")
            },
            Self::Execution(error) => error.fmt(formatter),
        }
    }
}

impl StdError for EvaluationFailure {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Allocation { source, .. } => Some(source),
            Self::ResourceLimit(limit) => Some(limit),
            Self::Execution(error) => Some(error),
            Self::Unsupported(_) | Self::InvalidExpression(_) | Self::Cancelled => None,
        }
    }
}

/// Result returned by scalar evaluation.
pub type EvaluationResult<T> = Result<T, EvaluationFailure>;

/// Evaluate one already-parsed ODS expression using the caller's execution
/// policy and finite evaluation limits.
pub fn evaluate_scalar<'a>(
    expression: &'a Expression,
    context: &EvaluationContext<'_>,
    limits: &EvaluationLimits,
) -> EvaluationResult<EvaluatedScalar<'a>> {
    context.execution.check().map_err(map_execution_error)?;
    let storage_budget = context.execution.budget().child(
        EVALUATION_SCOPE,
        litchi_core::Limits::new(
            u64::try_from(limits.max_storage_bytes).unwrap_or(u64::MAX),
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let mut evaluator = Evaluator::new(expression, context, *limits, storage_budget);
    let value = evaluator.run()?;
    Ok(evaluator.into_public(value))
}

/// Evaluate with explicit options without constructing an intermediate
/// [`EvaluationContext`].
pub fn evaluate_scalar_with_options<'a>(
    expression: &'a Expression,
    execution: &ExecutionContext,
    options: EvaluationOptions,
    limits: &EvaluationLimits,
) -> EvaluationResult<EvaluatedScalar<'a>> {
    let context = EvaluationContext::with_options(execution, options);
    evaluate_scalar(expression, &context, limits)
}

/// Evaluate using deterministic default options.
pub fn evaluate_scalar_with_context<'a>(
    expression: &'a Expression,
    execution: &ExecutionContext,
    limits: &EvaluationLimits,
) -> EvaluationResult<EvaluatedScalar<'a>> {
    evaluate_scalar(expression, &EvaluationContext::new(execution), limits)
}

struct Evaluator<'a, 'ctx, 'exec> {
    expression: &'a Expression,
    context: &'ctx EvaluationContext<'exec>,
    limits: EvaluationLimits,
    storage_budget: Budget,
    steps: u64,
    frames: Vec<Frame<'a>>,
    values: Vec<WorkingValue<'a>>,
    frame_reservation: Option<Reservation>,
    value_reservation: Option<Reservation>,
}

#[derive(Clone, Copy)]
enum Frame<'a> {
    Visit(Node<'a>),
    Apply(Node<'a>),
    VisitArgument(Node<'a>),
}

enum WorkingValue<'a> {
    Number(f64),
    Logical(bool),
    Text(TextValue<'a>),
    Error(ScalarError),
}

struct TextValue<'a> {
    text: Cow<'a, str>,
    reservation: Option<Reservation>,
}

struct NumberText {
    bytes: [u8; MAX_NUMBER_TEXT_BYTES],
    length: usize,
}

struct StringScan {
    decoded_len: usize,
    has_escaped_quotes: bool,
}

impl NumberText {
    fn new() -> Self {
        Self {
            bytes: [0; MAX_NUMBER_TEXT_BYTES],
            length: 0,
        }
    }

    fn as_str(&self) -> Result<&str, fmt::Error> {
        std::str::from_utf8(&self.bytes[..self.length]).map_err(|_| fmt::Error)
    }
}

impl FmtWrite for NumberText {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        let end = self.length.checked_add(value.len()).ok_or(fmt::Error)?;
        if end > self.bytes.len() {
            return Err(fmt::Error);
        }
        self.bytes[self.length..end].copy_from_slice(value.as_bytes());
        self.length = end;
        Ok(())
    }
}

fn copy_string_segment(
    evaluator: &mut Evaluator<'_, '_, '_>,
    output: &mut String,
    segment: &str,
) -> EvaluationResult<()> {
    let mut offset = 0usize;
    while offset < segment.len() {
        evaluator
            .context
            .execution
            .check()
            .map_err(map_execution_error)?;
        let mut end = offset
            .saturating_add(STRING_COPY_CHUNK_BYTES)
            .min(segment.len());
        while end < segment.len() && !segment.is_char_boundary(end) {
            end += 1;
        }
        let chunk = &segment[offset..end];
        evaluator.charge_bytes(chunk.len())?;
        output.push_str(chunk);
        offset = end;
    }
    Ok(())
}

impl<'a> TextValue<'a> {
    fn borrowed(text: &'a str) -> Self {
        Self {
            text: Cow::Borrowed(text),
            reservation: None,
        }
    }

    fn owned(text: String, reservation: Reservation) -> Self {
        Self {
            text: Cow::Owned(text),
            reservation: Some(reservation),
        }
    }
}

impl<'a, 'ctx, 'exec> Evaluator<'a, 'ctx, 'exec> {
    fn new(
        expression: &'a Expression,
        context: &'ctx EvaluationContext<'exec>,
        limits: EvaluationLimits,
        storage_budget: Budget,
    ) -> Self {
        Self {
            expression,
            context,
            limits,
            storage_budget,
            steps: 0,
            frames: Vec::new(),
            values: Vec::new(),
            frame_reservation: None,
            value_reservation: None,
        }
    }

    fn run(&mut self) -> EvaluationResult<WorkingValue<'a>> {
        let root = self.expression.root();
        self.push_frame(Frame::Visit(root))?;

        while let Some(frame) = self.frames.pop() {
            self.step()?;
            match frame {
                Frame::Visit(node) => self.visit(node)?,
                Frame::Apply(node) => self.apply(node)?,
                Frame::VisitArgument(node) => self.visit_argument(node)?,
            }
        }

        if self.values.len() != 1 {
            return Err(EvaluationFailure::InvalidExpression(
                "scalar evaluation did not produce exactly one value",
            ));
        }
        self.values
            .pop()
            .ok_or(EvaluationFailure::InvalidExpression(
                "scalar evaluation value stack was empty",
            ))
    }

    fn step(&mut self) -> EvaluationResult<()> {
        self.charge_work(1)
    }

    fn charge_work(&mut self, amount: u64) -> EvaluationResult<()> {
        self.context
            .execution
            .check()
            .map_err(map_execution_error)?;
        let next = self
            .steps
            .checked_add(amount)
            .ok_or_else(|| local_limit(Resource::Work, u64::MAX, self.limits.max_steps))?;
        if next > self.limits.max_steps {
            return Err(local_limit(Resource::Work, next, self.limits.max_steps));
        }
        if amount != 0 {
            self.context
                .execution
                .consume(Resource::Work, amount)
                .map_err(map_execution_error)?;
        }
        self.steps = next;
        Ok(())
    }

    fn charge_bytes(&mut self, bytes: usize) -> EvaluationResult<()> {
        self.charge_work(u64::try_from(bytes).unwrap_or(u64::MAX))
    }

    fn visit(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        match node.kind() {
            Kind::Number => {
                let value = self.parse_number(node.text())?;
                self.push_value(value)
            },
            Kind::String => {
                let value = self.parse_string(node.text())?;
                self.push_value(value)
            },
            Kind::Error => {
                let text = node.text();
                self.charge_bytes(text.len())?;
                self.push_value(WorkingValue::Error(parse_error(text)))
            },
            Kind::Parenthesized => self.push_frame(Frame::Visit(node.child(0).ok_or(
                EvaluationFailure::InvalidExpression("parenthesized node has no child"),
            )?)),
            Kind::Prefix(_) => {
                self.push_frame(Frame::Apply(node))?;
                self.push_frame(Frame::Visit(node.child(0).ok_or(
                    EvaluationFailure::InvalidExpression("prefix node has no child"),
                )?))
            },
            Kind::Postfix(_) => {
                self.push_frame(Frame::Apply(node))?;
                self.push_frame(Frame::Visit(node.child(0).ok_or(
                    EvaluationFailure::InvalidExpression("postfix node has no child"),
                )?))
            },
            Kind::Infix(operator) => {
                if matches!(
                    operator,
                    InfixOperator::Range | InfixOperator::Intersection | InfixOperator::Union
                ) {
                    return Err(EvaluationFailure::Unsupported(
                        UnsupportedKind::ReferenceOperator,
                    ));
                }
                let left = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                    "infix node has no left child",
                ))?;
                let right = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                    "infix node has no right child",
                ))?;
                self.push_frame(Frame::Apply(node))?;
                self.push_frame(Frame::Visit(right))?;
                self.push_frame(Frame::Visit(left))
            },
            Kind::Function { name } => self.visit_function(node, name),
            Kind::Reference(_) => Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference)),
            Kind::Array(_) | Kind::ArrayRow => {
                Err(EvaluationFailure::Unsupported(UnsupportedKind::Array))
            },
            Kind::NamedExpression { .. } => Err(EvaluationFailure::Unsupported(
                UnsupportedKind::NamedExpression,
            )),
            Kind::QuotedLabel | Kind::AutomaticIntersection => {
                Err(EvaluationFailure::Unsupported(UnsupportedKind::Label))
            },
            Kind::Missing => Err(EvaluationFailure::Unsupported(
                UnsupportedKind::MissingArgument,
            )),
        }
    }

    fn visit_argument(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        if node.is_missing() {
            self.push_value(WorkingValue::Error(ScalarError::Value))
        } else {
            self.visit(node)
        }
    }

    fn visit_function(&mut self, node: Node<'a>, name: &'a str) -> EvaluationResult<()> {
        self.charge_bytes(name.len())?;

        if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
            if node.child_count() == 0 {
                return self.push_value(WorkingValue::Logical(name.eq_ignore_ascii_case("TRUE")));
            }
            return self.schedule_eager_function(node);
        }

        if name.eq_ignore_ascii_case("IF") {
            let count = node.child_count();
            if !(1..=3).contains(&count) {
                return self.push_value(WorkingValue::Error(ScalarError::Value));
            }
            let condition = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                "IF condition is missing from the expression tree",
            ))?;
            self.push_frame(Frame::Apply(node))?;
            return self.push_frame(Frame::VisitArgument(condition));
        }

        if name.eq_ignore_ascii_case("IFERROR") || name.eq_ignore_ascii_case("IFNA") {
            if node.child_count() != 2 {
                return self.push_value(WorkingValue::Error(ScalarError::Value));
            }
            let value = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                "error-handling value is missing from the expression tree",
            ))?;
            self.push_frame(Frame::Apply(node))?;
            return self.push_frame(Frame::VisitArgument(value));
        }

        if name.eq_ignore_ascii_case("AND")
            || name.eq_ignore_ascii_case("OR")
            || name.eq_ignore_ascii_case("XOR")
            || name.eq_ignore_ascii_case("NOT")
            || name.eq_ignore_ascii_case("BITAND")
            || name.eq_ignore_ascii_case("BITLSHIFT")
            || name.eq_ignore_ascii_case("BITOR")
            || name.eq_ignore_ascii_case("BITRSHIFT")
            || name.eq_ignore_ascii_case("BITXOR")
        {
            return self.schedule_eager_function(node);
        }

        if radix::is_radix_function(name) {
            return self.schedule_eager_function(node);
        }

        Err(EvaluationFailure::Unsupported(UnsupportedKind::Function))
    }

    fn schedule_eager_function(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        let count = node.child_count();
        self.charge_work(u64::try_from(count).unwrap_or(u64::MAX))?;
        self.push_frame(Frame::Apply(node))?;
        let mut processed = 0usize;
        let mut next_check = 0usize;
        for index in (0..count).rev() {
            if processed >= next_check {
                self.context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                next_check = processed.saturating_add(4096);
            }
            processed = processed.saturating_add(1);
            let child = node
                .child(index)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "function argument is missing from the expression tree",
                ))?;
            self.push_frame(Frame::VisitArgument(child))?;
        }
        Ok(())
    }

    fn apply(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        match node.kind() {
            Kind::Function { name } => self.apply_function(node, name),
            Kind::Prefix(operator) => {
                let value = self.pop_value()?;
                let value = apply_prefix(operator, value, self)?;
                self.push_value(value)
            },
            Kind::Postfix(operator) => {
                let value = self.pop_value()?;
                let value = apply_postfix(operator, value, self)?;
                self.push_value(value)
            },
            Kind::Infix(operator) => {
                let right = self.pop_value()?;
                let left = self.pop_value()?;
                let value = self.apply_infix(operator, left, right)?;
                self.push_value(value)
            },
            _ => Err(EvaluationFailure::InvalidExpression(
                "non-operator reached evaluator apply frame",
            )),
        }
    }

    fn apply_function(&mut self, node: Node<'a>, name: &str) -> EvaluationResult<()> {
        if name.eq_ignore_ascii_case("IF") {
            return self.dispatch_if(node);
        }
        if name.eq_ignore_ascii_case("IFERROR") || name.eq_ignore_ascii_case("IFNA") {
            return self.dispatch_if_error(node);
        }
        if name.eq_ignore_ascii_case("AND") || name.eq_ignore_ascii_case("OR") {
            return self.apply_and_or(node, name.eq_ignore_ascii_case("AND"));
        }
        if name.eq_ignore_ascii_case("XOR") {
            return self.apply_xor(node);
        }
        if name.eq_ignore_ascii_case("NOT") {
            return self.apply_not(node);
        }
        if name.eq_ignore_ascii_case("BITAND")
            || name.eq_ignore_ascii_case("BITLSHIFT")
            || name.eq_ignore_ascii_case("BITOR")
            || name.eq_ignore_ascii_case("BITRSHIFT")
            || name.eq_ignore_ascii_case("BITXOR")
        {
            return self.apply_bitwise(node, name);
        }
        if radix::is_radix_function(name) {
            return radix::apply(self, node, name);
        }

        // TRUE/FALSE reach this path only for an invalid arity.  Consume all
        // scheduled arguments so the value stack remains balanced and retain
        // the leftmost formula error if one was produced.
        self.finish_invalid_arity(node)
    }

    #[inline(never)]
    fn apply_bitwise(&mut self, node: Node<'a>, name: &str) -> EvaluationResult<()> {
        if node.child_count() != 2 {
            return self.finish_invalid_arity(node);
        }

        let right = self.pop_value()?;
        let left = self.pop_value()?;
        let value = if name.eq_ignore_ascii_case("BITAND") {
            bitwise_pair(self, left, right, |left, right| left & right)?
        } else if name.eq_ignore_ascii_case("BITOR") {
            bitwise_pair(self, left, right, |left, right| left | right)?
        } else if name.eq_ignore_ascii_case("BITXOR") {
            bitwise_pair(self, left, right, |left, right| left ^ right)?
        } else if name.eq_ignore_ascii_case("BITLSHIFT") {
            bitwise_shift(self, left, right, true)?
        } else if name.eq_ignore_ascii_case("BITRSHIFT") {
            bitwise_shift(self, left, right, false)?
        } else {
            return Err(EvaluationFailure::InvalidExpression(
                "unknown bitwise function reached evaluator",
            ));
        };
        self.push_value(value)
    }

    fn finish_invalid_arity(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        let mut propagated = None;
        let mut next_check = 0usize;
        for index in 0..node.child_count() {
            if index >= next_check {
                self.context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                next_check = index.saturating_add(4096);
            }
            if let WorkingValue::Error(error) = self.pop_value()? {
                propagated = Some(error);
            }
        }
        self.push_value(WorkingValue::Error(
            propagated.unwrap_or(ScalarError::Value),
        ))
    }

    fn apply_and_or(&mut self, node: Node<'a>, conjunction: bool) -> EvaluationResult<()> {
        if node.child_count() == 0 {
            return self.push_value(WorkingValue::Error(ScalarError::Value));
        }

        let mut result = conjunction;
        let mut propagated = None;
        let mut next_check = 0usize;
        for index in 0..node.child_count() {
            if index >= next_check {
                self.context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                next_check = index.saturating_add(4096);
            }
            let value = self.pop_value()?;
            match to_number_sequence(value, self)? {
                Ok(value) => {
                    if conjunction {
                        result &= value;
                    } else {
                        result |= value;
                    }
                },
                Err(error) => propagated = Some(error),
            }
        }
        self.push_value(match propagated {
            Some(error) => WorkingValue::Error(error),
            None => WorkingValue::Logical(result),
        })
    }

    fn apply_xor(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        if node.child_count() == 0 {
            return self.push_value(WorkingValue::Error(ScalarError::Value));
        }

        let mut result = false;
        let mut propagated = None;
        let mut next_check = 0usize;
        for index in 0..node.child_count() {
            if index >= next_check {
                self.context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                next_check = index.saturating_add(4096);
            }
            let value = self.pop_value()?;
            match to_logical(value, self)? {
                Ok(value) => result ^= value,
                Err(error) => propagated = Some(error),
            }
        }
        self.push_value(match propagated {
            Some(error) => WorkingValue::Error(error),
            None => WorkingValue::Logical(result),
        })
    }

    fn apply_not(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        if node.child_count() != 1 {
            let mut propagated = None;
            let mut next_check = 0usize;
            for index in 0..node.child_count() {
                if index >= next_check {
                    self.context
                        .execution
                        .check()
                        .map_err(map_execution_error)?;
                    next_check = index.saturating_add(4096);
                }
                if let WorkingValue::Error(error) = self.pop_value()? {
                    propagated = Some(error);
                }
            }
            return self.push_value(WorkingValue::Error(
                propagated.unwrap_or(ScalarError::Value),
            ));
        }

        let value = self.pop_value()?;
        let value = match to_logical(value, self)? {
            Ok(value) => WorkingValue::Logical(!value),
            Err(error) => WorkingValue::Error(error),
        };
        self.push_value(value)
    }

    fn dispatch_if(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        let condition = self.pop_value()?;
        let condition = match to_logical(condition, self)? {
            Ok(value) => value,
            Err(error) => return self.push_value(WorkingValue::Error(error)),
        };

        match node.child_count() {
            1 => self.push_value(WorkingValue::Logical(condition)),
            2 if !condition => self.push_value(WorkingValue::Logical(false)),
            2 => self.push_if_branch(node.child(1)),
            3 if condition => self.push_if_branch(node.child(1)),
            3 => self.push_if_branch(node.child(2)),
            _ => self.push_value(WorkingValue::Error(ScalarError::Value)),
        }
    }

    fn push_if_branch(&mut self, branch: Option<Node<'a>>) -> EvaluationResult<()> {
        let branch = branch.ok_or(EvaluationFailure::InvalidExpression(
            "IF branch is missing from the expression tree",
        ))?;
        if branch.is_missing() {
            self.push_value(WorkingValue::Number(0.0))
        } else {
            self.push_frame(Frame::Visit(branch))
        }
    }

    fn dispatch_if_error(&mut self, node: Node<'a>) -> EvaluationResult<()> {
        let value = self.pop_value()?;
        let catches = match &value {
            WorkingValue::Error(error) => match node.kind() {
                Kind::Function { name } if name.eq_ignore_ascii_case("IFERROR") => true,
                Kind::Function { name } if name.eq_ignore_ascii_case("IFNA") => {
                    *error == ScalarError::NotAvailable
                },
                _ => false,
            },
            _ => false,
        };

        if !catches {
            return self.push_value(value);
        }
        let alternative = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
            "error-handling alternative is missing from the expression tree",
        ))?;
        self.push_frame(Frame::VisitArgument(alternative))
    }

    fn apply_infix(
        &mut self,
        operator: InfixOperator,
        left: WorkingValue<'a>,
        right: WorkingValue<'a>,
    ) -> EvaluationResult<WorkingValue<'a>> {
        if let WorkingValue::Error(error) = &left {
            return Ok(WorkingValue::Error(*error));
        }
        if let WorkingValue::Error(error) = &right {
            return Ok(WorkingValue::Error(*error));
        }

        match operator {
            InfixOperator::Add => numeric_binary(self, left, right, |a, b| a + b),
            InfixOperator::Subtract => numeric_binary(self, left, right, |a, b| a - b),
            InfixOperator::Multiply => numeric_binary(self, left, right, |a, b| a * b),
            InfixOperator::Divide => match numeric_pair(self, left, right)? {
                Err(error) => Ok(WorkingValue::Error(error)),
                Ok((_, 0.0)) => Ok(WorkingValue::Error(ScalarError::DivisionByZero)),
                Ok((left, right)) => Ok(finite_number(left / right)),
            },
            InfixOperator::Power => match numeric_pair(self, left, right)? {
                Err(error) => Ok(WorkingValue::Error(error)),
                Ok((left, right)) if left == 0.0 && right == 0.0 => Ok(WorkingValue::Number(1.0)),
                Ok((left, right)) => Ok(finite_number(left.powf(right))),
            },
            InfixOperator::Concatenate => concatenate(self, left, right),
            InfixOperator::Equal => Ok(WorkingValue::Logical(compare_equal(
                self,
                &left,
                &right,
                self.context.options.text_case,
            )?)),
            InfixOperator::NotEqual => Ok(WorkingValue::Logical(!compare_equal(
                self,
                &left,
                &right,
                self.context.options.text_case,
            )?)),
            InfixOperator::Less
            | InfixOperator::LessEqual
            | InfixOperator::Greater
            | InfixOperator::GreaterEqual => {
                ordered_compare(self, operator, left, right, self.context.options.text_case)
            },
            InfixOperator::Range | InfixOperator::Intersection | InfixOperator::Union => Err(
                EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator),
            ),
        }
    }

    fn parse_number(&mut self, text: &'a str) -> EvaluationResult<WorkingValue<'a>> {
        self.charge_bytes(text.len())?;
        let value = fast_float2::parse::<f64, _>(text).unwrap_or(f64::NAN);
        Ok(if value.is_finite() {
            WorkingValue::Number(value)
        } else {
            WorkingValue::Error(ScalarError::Number)
        })
    }

    fn parse_string(&mut self, text: &'a str) -> EvaluationResult<WorkingValue<'a>> {
        let body = text
            .strip_prefix('"')
            .and_then(|value| value.strip_suffix('"'))
            .ok_or(EvaluationFailure::InvalidExpression(
                "string node is missing quote delimiters",
            ))?;
        let scan = self.scan_string_body(body)?;
        if !scan.has_escaped_quotes {
            if body.len() > self.limits.max_text_bytes {
                return Err(local_limit(
                    Resource::Memory,
                    u64::try_from(body.len()).unwrap_or(u64::MAX),
                    u64::try_from(self.limits.max_text_bytes).unwrap_or(u64::MAX),
                ));
            }
            return Ok(WorkingValue::Text(TextValue::borrowed(body)));
        }

        let decoded_len = scan.decoded_len;
        if decoded_len > self.limits.max_text_bytes {
            return Err(local_limit(
                Resource::Memory,
                u64::try_from(decoded_len).unwrap_or(u64::MAX),
                u64::try_from(self.limits.max_text_bytes).unwrap_or(u64::MAX),
            ));
        }

        let reservation = self.reserve_storage(decoded_len, "formula scalar text")?;
        let mut decoded = String::new();
        if decoded_len != 0 {
            decoded.try_reserve_exact(decoded_len).map_err(|source| {
                EvaluationFailure::Allocation {
                    resource: "formula scalar text",
                    source,
                }
            })?;
        }
        self.copy_decoded_string(body, &mut decoded)?;
        if decoded.len() != decoded_len {
            return Err(EvaluationFailure::InvalidExpression(
                "decoded string length did not match its admitted size",
            ));
        }
        Ok(WorkingValue::Text(TextValue::owned(decoded, reservation)))
    }

    fn scan_string_body(&mut self, body: &str) -> EvaluationResult<StringScan> {
        self.charge_bytes(body.len())?;
        let bytes = body.as_bytes();
        let mut decoded_len = body.len();
        let mut escaped_quotes = 0usize;
        let mut cursor = 0usize;
        let mut next_check = 0usize;
        while cursor < bytes.len() {
            if cursor >= next_check {
                self.context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                next_check = cursor.saturating_add(4096);
            }
            match bytes[cursor] {
                0 => {
                    return Err(EvaluationFailure::InvalidExpression(
                        "string literal contains NUL",
                    ));
                },
                b'"' => {
                    if bytes.get(cursor + 1) != Some(&b'"') {
                        return Err(EvaluationFailure::InvalidExpression(
                            "unpaired quote in string literal",
                        ));
                    }
                    escaped_quotes = escaped_quotes.checked_add(1).ok_or(
                        EvaluationFailure::InvalidExpression("string escape count overflow"),
                    )?;
                    decoded_len =
                        decoded_len
                            .checked_sub(1)
                            .ok_or(EvaluationFailure::InvalidExpression(
                                "string escape length underflow",
                            ))?;
                    cursor = cursor
                        .checked_add(2)
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "string escape offset overflow",
                        ))?;
                },
                _ => cursor += 1,
            }
        }
        Ok(StringScan {
            decoded_len,
            has_escaped_quotes: escaped_quotes != 0,
        })
    }

    fn copy_decoded_string(&mut self, body: &str, decoded: &mut String) -> EvaluationResult<()> {
        // Finding copy segments is a second source scan, separate from the
        // decoded bytes charged by copy_string_segment and quote writes.
        self.charge_bytes(body.len())?;
        let bytes = body.as_bytes();
        let mut segment_start = 0usize;
        let mut cursor = 0usize;
        let mut next_check = 0usize;
        while cursor < bytes.len() {
            if cursor >= next_check {
                self.context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                next_check = cursor.saturating_add(STRING_COPY_CHUNK_BYTES);
            }
            if bytes[cursor] == b'"' {
                if segment_start < cursor {
                    copy_string_segment(self, decoded, &body[segment_start..cursor])?;
                }
                self.charge_bytes(1)?;
                decoded.push('"');
                cursor = cursor
                    .checked_add(2)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "string escape offset overflow",
                    ))?;
                segment_start = cursor;
            } else {
                cursor += 1;
            }
        }
        if segment_start < body.len() {
            copy_string_segment(self, decoded, &body[segment_start..])?;
        }
        Ok(())
    }

    fn reserve_storage(&self, bytes: usize, _scope: &'static str) -> EvaluationResult<Reservation> {
        if bytes > self.limits.max_storage_bytes {
            return Err(local_limit(
                Resource::Memory,
                u64::try_from(bytes).unwrap_or(u64::MAX),
                u64::try_from(self.limits.max_storage_bytes).unwrap_or(u64::MAX),
            ));
        }
        self.context
            .execution
            .check()
            .map_err(map_execution_error)?;
        self.storage_budget
            .reserve(Resource::Memory, u64::try_from(bytes).unwrap_or(u64::MAX))
            .map_err(EvaluationFailure::ResourceLimit)
    }

    fn push_frame(&mut self, frame: Frame<'a>) -> EvaluationResult<()> {
        self.ensure_frame_capacity(1)?;
        self.frames.push(frame);
        Ok(())
    }

    fn push_value(&mut self, value: WorkingValue<'a>) -> EvaluationResult<()> {
        self.ensure_value_capacity(1)?;
        self.values.push(value);
        Ok(())
    }

    fn pop_value(&mut self) -> EvaluationResult<WorkingValue<'a>> {
        self.values
            .pop()
            .ok_or(EvaluationFailure::InvalidExpression(
                "operator value stack underflow",
            ))
    }

    fn ensure_frame_capacity(&mut self, additional: usize) -> EvaluationResult<()> {
        let execution = self.context.execution;
        let storage_budget = &self.storage_budget;
        ensure_capacity(
            &mut self.frames,
            &mut self.frame_reservation,
            additional,
            self.limits.max_stack_entries,
            execution,
            storage_budget,
            "formula scalar frame stack",
        )
    }

    fn ensure_value_capacity(&mut self, additional: usize) -> EvaluationResult<()> {
        let execution = self.context.execution;
        let storage_budget = &self.storage_budget;
        ensure_capacity(
            &mut self.values,
            &mut self.value_reservation,
            additional,
            self.limits.max_stack_entries,
            execution,
            storage_budget,
            "formula scalar value stack",
        )
    }

    fn into_public(self, value: WorkingValue<'a>) -> EvaluatedScalar<'a> {
        match value {
            WorkingValue::Number(value) => EvaluatedScalar::new(ScalarValue::Number(value), None),
            WorkingValue::Logical(value) => EvaluatedScalar::new(ScalarValue::Logical(value), None),
            WorkingValue::Error(value) => EvaluatedScalar::new(ScalarValue::Error(value), None),
            WorkingValue::Text(value) => {
                EvaluatedScalar::new(ScalarValue::Text(value.text), value.reservation)
            },
        }
    }
}

fn ensure_capacity<T>(
    values: &mut Vec<T>,
    reservation: &mut Option<Reservation>,
    additional: usize,
    maximum: usize,
    execution: &ExecutionContext,
    storage_budget: &Budget,
    scope: &'static str,
) -> EvaluationResult<()> {
    let required =
        values
            .len()
            .checked_add(additional)
            .ok_or(EvaluationFailure::InvalidExpression(
                "evaluator stack length overflow",
            ))?;
    if required > maximum {
        return Err(local_limit(
            Resource::Objects,
            u64::try_from(required).unwrap_or(u64::MAX),
            u64::try_from(maximum).unwrap_or(u64::MAX),
        ));
    }
    if required <= values.capacity() {
        return Ok(());
    }
    let grown = values
        .capacity()
        .max(1)
        .checked_mul(2)
        .unwrap_or(maximum)
        .max(required)
        .min(maximum);
    let bytes = grown
        .checked_mul(size_of::<T>())
        .ok_or(EvaluationFailure::InvalidExpression(
            "evaluator stack byte size overflow",
        ))?;
    execution.check().map_err(map_execution_error)?;
    let reservation_new = storage_budget
        .reserve(Resource::Memory, u64::try_from(bytes).unwrap_or(u64::MAX))
        .map_err(EvaluationFailure::ResourceLimit)?;
    let old_reservation = reservation.take();
    let result = values.try_reserve_exact(grown.saturating_sub(values.len()));
    if let Err(source) = result {
        *reservation = old_reservation;
        return Err(EvaluationFailure::Allocation {
            resource: scope,
            source,
        });
    }
    drop(old_reservation);
    *reservation = Some(reservation_new);
    Ok(())
}

fn apply_prefix<'a>(
    operator: PrefixOperator,
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<WorkingValue<'a>> {
    match operator {
        PrefixOperator::Plus => Ok(value),
        PrefixOperator::Minus => match to_number(value, evaluator)? {
            Ok(number) => Ok(finite_number(-number)),
            Err(error) => Ok(WorkingValue::Error(error)),
        },
    }
}

fn apply_postfix<'a>(
    operator: PostfixOperator,
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<WorkingValue<'a>> {
    match operator {
        PostfixOperator::Percent => match to_number(value, evaluator)? {
            Ok(number) => Ok(finite_number(number / 100.0)),
            Err(error) => Ok(WorkingValue::Error(error)),
        },
    }
}

fn numeric_binary<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    left: WorkingValue<'a>,
    right: WorkingValue<'a>,
    operation: impl FnOnce(f64, f64) -> f64,
) -> EvaluationResult<WorkingValue<'a>> {
    match numeric_pair(evaluator, left, right)? {
        Ok((left, right)) => Ok(finite_number(operation(left, right))),
        Err(error) => Ok(WorkingValue::Error(error)),
    }
}

fn numeric_pair<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    left: WorkingValue<'a>,
    right: WorkingValue<'a>,
) -> EvaluationResult<Result<(f64, f64), ScalarError>> {
    let left = match to_number(left, evaluator)? {
        Ok(value) => value,
        Err(error) => return Ok(Err(error)),
    };
    let right = match to_number(right, evaluator)? {
        Ok(value) => value,
        Err(error) => return Ok(Err(error)),
    };
    Ok(Ok((left, right)))
}

fn finite_number<'a>(value: f64) -> WorkingValue<'a> {
    if value.is_finite() {
        WorkingValue::Number(value)
    } else {
        WorkingValue::Error(ScalarError::Number)
    }
}

fn to_number<'a>(
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    match value {
        WorkingValue::Number(value) => Ok(Ok(value)),
        WorkingValue::Logical(value) => Ok(Ok(if value { 1.0 } else { 0.0 })),
        WorkingValue::Text(text) => {
            evaluator.charge_bytes(text.text.len())?;
            let parsed = fast_float2::parse::<f64, _>(text.text.as_ref()).ok();
            match parsed.filter(|value| value.is_finite()) {
                Some(value) => Ok(Ok(value)),
                None => Ok(Err(ScalarError::Value)),
            }
        },
        WorkingValue::Error(error) => Ok(Err(error)),
    }
}

const BIT_WIDTH: u32 = 48;
const BIT_MAX: u64 = (1_u64 << BIT_WIDTH) - 1;
const BIT_LIMIT: f64 = 281_474_976_710_656.0;

/// Apply the profile's Conversion to Integer rule: first use the existing
/// Number conversion and then truncate toward zero.  The latter is an
/// explicit profile choice because Part 4 leaves the generic conversion from
/// non-integer Numbers implementation-defined.
fn to_integer<'a>(
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    match to_number(value, evaluator)? {
        Ok(value) => {
            let value = value.trunc();
            if value.is_finite() {
                Ok(Ok(value))
            } else {
                Ok(Err(ScalarError::Number))
            }
        },
        Err(error) => Ok(Err(error)),
    }
}

/// Convert one integer operand to the supported unsigned 48-bit domain.
fn to_bit_operand<'a>(
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<u64, ScalarError>> {
    match to_integer(value, evaluator)? {
        Ok(value) if (0.0..BIT_LIMIT).contains(&value) => Ok(Ok(value as u64)),
        Ok(_) => Ok(Err(ScalarError::Number)),
        Err(error) => Ok(Err(error)),
    }
}

fn bitwise_pair<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    left: WorkingValue<'a>,
    right: WorkingValue<'a>,
    operation: impl FnOnce(u64, u64) -> u64,
) -> EvaluationResult<WorkingValue<'a>> {
    let left = to_bit_operand(left, evaluator)?;
    let right = to_bit_operand(right, evaluator)?;
    let result = match (left, right) {
        (Err(error), _) => return Ok(WorkingValue::Error(error)),
        (Ok(_), Err(error)) => return Ok(WorkingValue::Error(error)),
        (Ok(left), Ok(right)) => operation(left, right),
    };
    Ok(bit_result(result))
}

fn bitwise_shift<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    left: WorkingValue<'a>,
    right: WorkingValue<'a>,
    leftward: bool,
) -> EvaluationResult<WorkingValue<'a>> {
    let left = to_bit_operand(left, evaluator)?;
    let right = to_integer(right, evaluator)?;
    let left = match left {
        Ok(value) => value,
        Err(error) => return Ok(WorkingValue::Error(error)),
    };
    let right = match right {
        Ok(value) => value,
        Err(error) => return Ok(WorkingValue::Error(error)),
    };

    if right < 0.0 {
        return Ok(if leftward {
            shift_right(left, -right)
        } else {
            shift_left(left, -right)
        });
    }
    Ok(if leftward {
        shift_left(left, right)
    } else {
        shift_right(left, right)
    })
}

fn bit_result<'a>(value: u64) -> WorkingValue<'a> {
    if value <= BIT_MAX {
        WorkingValue::Number(value as f64)
    } else {
        WorkingValue::Error(ScalarError::Number)
    }
}

fn shift_left<'a>(value: u64, amount: f64) -> WorkingValue<'a> {
    if value == 0 || amount == 0.0 {
        return WorkingValue::Number(value as f64);
    }
    if amount >= f64::from(BIT_WIDTH) {
        return WorkingValue::Error(ScalarError::Number);
    }
    let amount = amount as u32;
    // `checked_shl` only checks the shift count; it may still discard high
    // bits when the shifted value exceeds the native integer width.  Check
    // the profile's 48-bit result domain before performing the shift.
    if value > (BIT_MAX >> amount) {
        return WorkingValue::Error(ScalarError::Number);
    }
    match value.checked_shl(amount) {
        Some(value) => bit_result(value),
        None => WorkingValue::Error(ScalarError::Number),
    }
}

fn shift_right<'a>(value: u64, amount: f64) -> WorkingValue<'a> {
    if amount >= f64::from(BIT_WIDTH) {
        return WorkingValue::Number(0.0);
    }
    if amount == 0.0 {
        return WorkingValue::Number(value as f64);
    }
    WorkingValue::Number((value >> amount as u32) as f64)
}

/// Conversion used by the scalar Logical parameter family.  Text conversion
/// is deliberately deterministic for this profile: text is not accepted as a
/// logical spelling and produces `#VALUE!`.  AND/OR use the separate numeric
/// sequence conversion below, as required by their OpenFormula signatures.
fn to_logical<'a>(
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<bool, ScalarError>> {
    match value {
        WorkingValue::Number(value) => Ok(Ok(value != 0.0)),
        WorkingValue::Logical(value) => Ok(Ok(value)),
        WorkingValue::Text(text) => {
            evaluator.charge_bytes(text.text.len())?;
            Ok(Err(ScalarError::Value))
        },
        WorkingValue::Error(error) => Ok(Err(error)),
    }
}

/// Conversion for the `NumberSequenceList` alternative accepted by AND and
/// OR.  A scalar Text therefore follows Conversion to Number, while the
/// other Logical functions use `to_logical` directly.
fn to_number_sequence<'a>(
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<Result<bool, ScalarError>> {
    match to_number(value, evaluator)? {
        Ok(value) => Ok(Ok(value != 0.0)),
        Err(error) => Ok(Err(error)),
    }
}

fn to_text<'a>(
    value: WorkingValue<'a>,
    evaluator: &mut Evaluator<'_, '_, '_>,
) -> EvaluationResult<TextValue<'a>> {
    match value {
        WorkingValue::Text(value) => Ok(value),
        WorkingValue::Logical(value) => {
            Ok(TextValue::borrowed(if value { "TRUE" } else { "FALSE" }))
        },
        WorkingValue::Number(value) => {
            evaluator.charge_work(1)?;
            let mut number = NumberText::new();
            write!(&mut number, "{value}").map_err(|_| {
                EvaluationFailure::InvalidExpression("number-to-text formatting failed")
            })?;
            let rendered = number
                .as_str()
                .map_err(|_| EvaluationFailure::InvalidExpression("invalid formatted number"))?;
            evaluator.charge_bytes(rendered.len())?;
            if rendered.len() > evaluator.limits.max_text_bytes {
                return Err(local_limit(
                    Resource::Memory,
                    u64::try_from(rendered.len()).unwrap_or(u64::MAX),
                    u64::try_from(evaluator.limits.max_text_bytes).unwrap_or(u64::MAX),
                ));
            }
            let reservation =
                evaluator.reserve_storage(rendered.len(), "formula scalar number text")?;
            let mut text = String::new();
            if !rendered.is_empty() {
                text.try_reserve_exact(rendered.len()).map_err(|source| {
                    EvaluationFailure::Allocation {
                        resource: "formula scalar number text",
                        source,
                    }
                })?;
            }
            text.push_str(rendered);
            Ok(TextValue::owned(text, reservation))
        },
        WorkingValue::Error(_) => Err(EvaluationFailure::InvalidExpression(
            "error reached text conversion after propagation check",
        )),
    }
}

fn concatenate<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    left: WorkingValue<'a>,
    right: WorkingValue<'a>,
) -> EvaluationResult<WorkingValue<'a>> {
    let left = to_text(left, evaluator)?;
    let right = to_text(right, evaluator)?;
    let left_len = left.text.len();
    let right_len = right.text.len();
    let total = left_len.checked_add(right_len).ok_or_else(|| {
        local_limit(
            Resource::Memory,
            u64::MAX,
            evaluator.limits.max_text_bytes as u64,
        )
    })?;
    if total > evaluator.limits.max_text_bytes {
        return Err(local_limit(
            Resource::Memory,
            u64::try_from(total).unwrap_or(u64::MAX),
            u64::try_from(evaluator.limits.max_text_bytes).unwrap_or(u64::MAX),
        ));
    }

    if left_len == 0 {
        return Ok(WorkingValue::Text(right));
    }
    if right_len == 0 {
        return Ok(WorkingValue::Text(left));
    }
    let work = match &left.text {
        Cow::Owned(output) if output.capacity() >= total => right_len,
        _ => total,
    };
    evaluator.charge_bytes(work)?;

    match left.text {
        Cow::Owned(mut output) => {
            let current_capacity = output.capacity();
            let mut output_reservation = left.reservation;
            if current_capacity < total {
                let target_capacity = current_capacity
                    .max(1)
                    .checked_mul(2)
                    .unwrap_or(total)
                    .max(total);
                let reservation =
                    evaluator.reserve_storage(target_capacity, "formula scalar concatenation")?;
                output
                    .try_reserve_exact(target_capacity.saturating_sub(output.len()))
                    .map_err(|source| EvaluationFailure::Allocation {
                        resource: "formula scalar concatenation",
                        source,
                    })?;
                drop(output_reservation.take());
                output_reservation = Some(reservation);
            }
            output.push_str(right.text.as_ref());
            drop(right.reservation);
            Ok(WorkingValue::Text(TextValue {
                text: Cow::Owned(output),
                reservation: output_reservation,
            }))
        },
        Cow::Borrowed(left_text) => {
            let reservation = evaluator.reserve_storage(total, "formula scalar concatenation")?;
            let mut output = String::new();
            if total != 0 {
                output.try_reserve_exact(total).map_err(|source| {
                    EvaluationFailure::Allocation {
                        resource: "formula scalar concatenation",
                        source,
                    }
                })?;
            }
            output.push_str(left_text);
            output.push_str(right.text.as_ref());
            drop(right.reservation);
            Ok(WorkingValue::Text(TextValue::owned(output, reservation)))
        },
    }
}

fn compare_equal(
    evaluator: &mut Evaluator<'_, '_, '_>,
    left: &WorkingValue<'_>,
    right: &WorkingValue<'_>,
    case: TextCase,
) -> EvaluationResult<bool> {
    match (left, right) {
        (WorkingValue::Number(left), WorkingValue::Number(right)) => Ok(left == right),
        (WorkingValue::Logical(left), WorkingValue::Logical(right)) => Ok(left == right),
        (WorkingValue::Text(left), WorkingValue::Text(right)) => {
            let bytes = left
                .text
                .len()
                .checked_add(right.text.len())
                .ok_or_else(|| local_limit(Resource::Work, u64::MAX, evaluator.limits.max_steps))?;
            evaluator.charge_bytes(bytes)?;
            Ok(
                compare_text(evaluator, left.text.as_ref(), right.text.as_ref(), case)?
                    == Ordering::Equal,
            )
        },
        _ => Ok(false),
    }
}

fn ordered_compare<'a>(
    evaluator: &mut Evaluator<'_, '_, '_>,
    operator: InfixOperator,
    left: WorkingValue<'a>,
    right: WorkingValue<'a>,
    case: TextCase,
) -> EvaluationResult<WorkingValue<'a>> {
    let ordering = match (&left, &right) {
        (WorkingValue::Number(left), WorkingValue::Number(right)) => left.partial_cmp(right),
        (WorkingValue::Logical(left), WorkingValue::Logical(right)) => {
            (*left as u8).partial_cmp(&(*right as u8))
        },
        (WorkingValue::Text(left), WorkingValue::Text(right)) => {
            let bytes = left
                .text
                .len()
                .checked_add(right.text.len())
                .ok_or_else(|| local_limit(Resource::Work, u64::MAX, evaluator.limits.max_steps))?;
            evaluator.charge_bytes(bytes)?;
            Some(compare_text(
                evaluator,
                left.text.as_ref(),
                right.text.as_ref(),
                case,
            )?)
        },
        _ => {
            return Ok(WorkingValue::Error(ScalarError::Value));
        },
    }
    .ok_or(EvaluationFailure::InvalidExpression(
        "non-finite number reached ordered comparison",
    ))?;

    let result = match operator {
        InfixOperator::Less => ordering == Ordering::Less,
        InfixOperator::LessEqual => ordering != Ordering::Greater,
        InfixOperator::Greater => ordering == Ordering::Greater,
        InfixOperator::GreaterEqual => ordering != Ordering::Less,
        _ => {
            return Err(EvaluationFailure::InvalidExpression(
                "non-ordering operator reached ordered comparison",
            ));
        },
    };
    Ok(WorkingValue::Logical(result))
}

fn compare_text(
    evaluator: &mut Evaluator<'_, '_, '_>,
    left: &str,
    right: &str,
    case: TextCase,
) -> EvaluationResult<Ordering> {
    match case {
        TextCase::Sensitive => {
            let shared = left.len().min(right.len());
            let mut offset = 0usize;
            while offset < shared {
                evaluator
                    .context
                    .execution
                    .check()
                    .map_err(map_execution_error)?;
                let end = offset.saturating_add(4096).min(shared);
                let ordering = left.as_bytes()[offset..end].cmp(&right.as_bytes()[offset..end]);
                if ordering != Ordering::Equal {
                    return Ok(ordering);
                }
                offset = end;
            }
            Ok(left.len().cmp(&right.len()))
        },
    }
}

fn parse_error(text: &str) -> ScalarError {
    if text.eq_ignore_ascii_case("#N/A") {
        ScalarError::NotAvailable
    } else if text.eq_ignore_ascii_case("#DIV/0!") {
        ScalarError::DivisionByZero
    } else if text.eq_ignore_ascii_case("#REF!") {
        ScalarError::Reference
    } else if text.eq_ignore_ascii_case("#NUM!") {
        ScalarError::Number
    } else if text.eq_ignore_ascii_case("#NULL!") {
        ScalarError::Null
    } else if text.eq_ignore_ascii_case("#VALUE!") {
        ScalarError::Value
    } else {
        ScalarError::Name
    }
}

fn local_limit(resource: Resource, observed: u64, limit: u64) -> EvaluationFailure {
    EvaluationFailure::ResourceLimit(ResourceLimit {
        resource,
        observed,
        limit,
        scope: Arc::from(EVALUATION_SCOPE),
    })
}

fn map_execution_error(error: ExecutionError) -> EvaluationFailure {
    match error {
        ExecutionError::Cancelled => EvaluationFailure::Cancelled,
        ExecutionError::ResourceLimit(limit) => EvaluationFailure::ResourceLimit(limit),
        other => EvaluationFailure::Execution(other),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Profile};
    use std::num::{NonZeroU64, NonZeroUsize};

    fn context() -> (CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            Arc::from("ods-formula-test"),
            litchi_core::Limits::for_profile(Profile::Desktop),
        );
        let (source, token) = CancellationSource::pair();
        let execution = ExecutionContext::new(
            budget,
            token,
            ExecutionLimits::new(
                NonZeroUsize::MIN,
                NonZeroUsize::MIN,
                NonZeroU64::new(1 << 30).unwrap(),
                1 << 20,
            )
            .unwrap(),
        );
        (source, execution)
    }

    #[test]
    fn evaluates_scalar_operators_without_recursive_walk() {
        let (_, execution) = context();
        for (source, expected) in [
            ("=2+3*4", 14.0),
            ("=-2^2", 4.0),
            ("=2^3^2", 64.0),
            ("=50%", 0.5),
        ] {
            let expression = Expression::parse(source).unwrap();
            let result =
                evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default())
                    .unwrap();
            assert_eq!(result.value(), &ScalarValue::Number(expected));
        }
    }

    #[test]
    fn preserves_text_and_maps_errors() {
        let (_, execution) = context();
        let expression = Expression::parse("=\"a\"&1").unwrap();
        let result =
            evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default())
                .unwrap();
        assert_eq!(
            result.value(),
            &ScalarValue::Text(Cow::Owned("a1".to_string()))
        );

        let expression = Expression::parse("=#N/A+1").unwrap();
        let result =
            evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default())
                .unwrap();
        assert_eq!(
            result.value(),
            &ScalarValue::Error(ScalarError::NotAvailable)
        );

        let expression = Expression::parse("=#UNKNOWN!").unwrap();
        let result =
            evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default())
                .unwrap();
        assert_eq!(result.value(), &ScalarValue::Error(ScalarError::Name));

        let expression = Expression::parse("=TRUE()&FALSE()").unwrap();
        let result =
            evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default())
                .unwrap();
        assert_eq!(
            result.value(),
            &ScalarValue::Text(Cow::Owned("TRUEFALSE".to_string()))
        );
    }

    #[test]
    fn rejects_capabilities_and_honors_cancellation() {
        let expression = Expression::parse("[.A1]").unwrap();
        let (_, execution) = context();
        let error =
            evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default())
                .unwrap_err();
        assert!(matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::Reference)
        ));

        let (source, execution) = context();
        source.cancel();
        let expression = Expression::parse("1+2").unwrap();
        assert!(matches!(
            evaluate_scalar_with_context(&expression, &execution, &EvaluationLimits::default()),
            Err(EvaluationFailure::Cancelled)
        ));
    }
}
