//! Typed `DrawingML` geometry-guide formulas (`a:gd@fmla`,
//! `ST_GeomGuideFormula`; ECMA-376 part 1 §20.1.10.11).
//!
//! A formula is an operation token followed by space-delimited operands,
//! each a literal or a guide reference. [`Formula`] models every
//! defined operation with its exact operand count; parsing rejects unknown
//! operations and wrong arities so authored formulas always serialize back
//! to a schema-valid token.

use std::fmt;
use std::str::FromStr;

use crate::{Error, Result};

use super::AdjustValue;

/// A geometry-guide formula: one ECMA-376 operation with typed operands.
///
/// Formula semantics quote §20.1.10.11; trigonometric operands are angles in
/// 60000ths of a degree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Formula {
    /// `*/ x y z` — `(x * y) / z`.
    MultiplyDivide {
        /// First factor.
        x: AdjustValue,
        /// Second factor.
        y: AdjustValue,
        /// Divisor.
        z: AdjustValue,
    },
    /// `+- x y z` — `(x + y) - z`.
    AddSubtract {
        /// First addend.
        x: AdjustValue,
        /// Second addend.
        y: AdjustValue,
        /// Subtrahend.
        z: AdjustValue,
    },
    /// `+/ x y z` — `(x + y) / z`.
    AddDivide {
        /// First addend.
        x: AdjustValue,
        /// Second addend.
        y: AdjustValue,
        /// Divisor.
        z: AdjustValue,
    },
    /// `?: x y z` — `y` when `x > 0`, `z` otherwise.
    IfElse {
        /// Condition value.
        x: AdjustValue,
        /// Result when the condition is positive.
        y: AdjustValue,
        /// Result otherwise.
        z: AdjustValue,
    },
    /// `abs x` — `|x|`.
    Absolute {
        /// Input value.
        x: AdjustValue,
    },
    /// `at2 x y` — `atan2(y, x)`.
    ArcTangent {
        /// Horizontal component.
        x: AdjustValue,
        /// Vertical component.
        y: AdjustValue,
    },
    /// `cat2 x y z` — `x * cos(atan2(z, y))`.
    CosineArcTangent {
        /// Scale value.
        x: AdjustValue,
        /// Horizontal component.
        y: AdjustValue,
        /// Vertical component.
        z: AdjustValue,
    },
    /// `cos x y` — `x * cos(y)`.
    Cosine {
        /// Scale value.
        x: AdjustValue,
        /// Angle.
        y: AdjustValue,
    },
    /// `max x y` — the greater of `x` and `y`.
    Maximum {
        /// First value.
        x: AdjustValue,
        /// Second value.
        y: AdjustValue,
    },
    /// `min x y` — the lesser of `x` and `y`.
    Minimum {
        /// First value.
        x: AdjustValue,
        /// Second value.
        y: AdjustValue,
    },
    /// `mod x y z` — `sqrt(x² + y² + z²)`.
    Modulus {
        /// First component.
        x: AdjustValue,
        /// Second component.
        y: AdjustValue,
        /// Third component.
        z: AdjustValue,
    },
    /// `pin x y z` — `y` clamped to `[x, z]`.
    Pin {
        /// Lower bound.
        x: AdjustValue,
        /// Clamped value.
        y: AdjustValue,
        /// Upper bound.
        z: AdjustValue,
    },
    /// `sat2 x y z` — `x * sin(atan2(z, y))`.
    SineArcTangent {
        /// Scale value.
        x: AdjustValue,
        /// Horizontal component.
        y: AdjustValue,
        /// Vertical component.
        z: AdjustValue,
    },
    /// `sin x y` — `x * sin(y)`.
    Sine {
        /// Scale value.
        x: AdjustValue,
        /// Angle.
        y: AdjustValue,
    },
    /// `sqrt x` — `sqrt(x)`.
    SquareRoot {
        /// Input value.
        x: AdjustValue,
    },
    /// `tan x y` — `x * tan(y)`.
    Tangent {
        /// Scale value.
        x: AdjustValue,
        /// Angle.
        y: AdjustValue,
    },
    /// `val x` — the literal or referenced value itself.
    Value {
        /// The value.
        x: AdjustValue,
    },
}

impl Formula {
    /// A `val` formula holding one literal, the common shape of adjust-value
    /// entries.
    #[must_use]
    pub fn literal(value: i64) -> Self {
        Self::Value {
            x: AdjustValue::Value(value),
        }
    }

    /// The formula's operands in serialization order.
    pub fn operands(&self) -> impl Iterator<Item = &AdjustValue> {
        match self {
            Self::MultiplyDivide { x, y, z }
            | Self::AddSubtract { x, y, z }
            | Self::AddDivide { x, y, z }
            | Self::IfElse { x, y, z }
            | Self::CosineArcTangent { x, y, z }
            | Self::Modulus { x, y, z }
            | Self::Pin { x, y, z }
            | Self::SineArcTangent { x, y, z } => FormulaOperands::three(x, y, z),
            Self::ArcTangent { x, y }
            | Self::Cosine { x, y }
            | Self::Maximum { x, y }
            | Self::Minimum { x, y }
            | Self::Sine { x, y }
            | Self::Tangent { x, y } => FormulaOperands::two(x, y),
            Self::Absolute { x } | Self::SquareRoot { x } | Self::Value { x } => {
                FormulaOperands::one(x)
            },
        }
    }

    /// The formula's `ST_GeomGuideFormula` operation token.
    #[must_use]
    pub fn operation(&self) -> &'static str {
        match self {
            Self::MultiplyDivide { .. } => "*/",
            Self::AddSubtract { .. } => "+-",
            Self::AddDivide { .. } => "+/",
            Self::IfElse { .. } => "?:",
            Self::Absolute { .. } => "abs",
            Self::ArcTangent { .. } => "at2",
            Self::CosineArcTangent { .. } => "cat2",
            Self::Cosine { .. } => "cos",
            Self::Maximum { .. } => "max",
            Self::Minimum { .. } => "min",
            Self::Modulus { .. } => "mod",
            Self::Pin { .. } => "pin",
            Self::SineArcTangent { .. } => "sat2",
            Self::Sine { .. } => "sin",
            Self::SquareRoot { .. } => "sqrt",
            Self::Tangent { .. } => "tan",
            Self::Value { .. } => "val",
        }
    }
}

/// Iterator over a formula's one to three operands.
struct FormulaOperands<'formula> {
    operands: [Option<&'formula AdjustValue>; 3],
    next: usize,
}

impl<'formula> FormulaOperands<'formula> {
    fn one(x: &'formula AdjustValue) -> Self {
        Self {
            operands: [Some(x), None, None],
            next: 0,
        }
    }

    fn two(x: &'formula AdjustValue, y: &'formula AdjustValue) -> Self {
        Self {
            operands: [Some(x), Some(y), None],
            next: 0,
        }
    }

    fn three(x: &'formula AdjustValue, y: &'formula AdjustValue, z: &'formula AdjustValue) -> Self {
        Self {
            operands: [Some(x), Some(y), Some(z)],
            next: 0,
        }
    }
}

impl<'formula> Iterator for FormulaOperands<'formula> {
    type Item = &'formula AdjustValue;

    fn next(&mut self) -> Option<Self::Item> {
        let operand = self.operands.get(self.next).copied().flatten()?;
        self.next += 1;
        Some(operand)
    }
}

impl fmt::Display for Formula {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.operation())?;
        for operand in self.operands() {
            write!(formatter, " {operand}")?;
        }
        Ok(())
    }
}

impl FromStr for Formula {
    type Err = Error;

    fn from_str(text: &str) -> Result<Self> {
        let mut tokens = text.split_whitespace();
        let operation = tokens
            .next()
            .ok_or_else(|| invalid_formula(text, "it has no operation"))?;
        let operands = tokens
            .map(AdjustValue::from_str)
            .collect::<Result<Vec<_>>>()?;
        let formula = match operation {
            "*/" => {
                let [x, y, z] = take_operands(text, operands)?;
                Self::MultiplyDivide { x, y, z }
            },
            "+-" => {
                let [x, y, z] = take_operands(text, operands)?;
                Self::AddSubtract { x, y, z }
            },
            "+/" => {
                let [x, y, z] = take_operands(text, operands)?;
                Self::AddDivide { x, y, z }
            },
            "?:" => {
                let [x, y, z] = take_operands(text, operands)?;
                Self::IfElse { x, y, z }
            },
            "abs" => {
                let [x] = take_operands(text, operands)?;
                Self::Absolute { x }
            },
            "at2" => {
                let [x, y] = take_operands(text, operands)?;
                Self::ArcTangent { x, y }
            },
            "cat2" => {
                let [x, y, z] = take_operands(text, operands)?;
                Self::CosineArcTangent { x, y, z }
            },
            "cos" => {
                let [x, y] = take_operands(text, operands)?;
                Self::Cosine { x, y }
            },
            "max" => {
                let [x, y] = take_operands(text, operands)?;
                Self::Maximum { x, y }
            },
            "min" => {
                let [x, y] = take_operands(text, operands)?;
                Self::Minimum { x, y }
            },
            "mod" => {
                let [x, y, z] = take_operands(text, operands)?;
                Self::Modulus { x, y, z }
            },
            "pin" => {
                let [x, y, z] = take_operands(text, operands)?;
                Self::Pin { x, y, z }
            },
            "sat2" => {
                let [x, y, z] = take_operands(text, operands)?;
                Self::SineArcTangent { x, y, z }
            },
            "sin" => {
                let [x, y] = take_operands(text, operands)?;
                Self::Sine { x, y }
            },
            "sqrt" => {
                let [x] = take_operands(text, operands)?;
                Self::SquareRoot { x }
            },
            "tan" => {
                let [x, y] = take_operands(text, operands)?;
                Self::Tangent { x, y }
            },
            "val" => {
                let [x] = take_operands(text, operands)?;
                Self::Value { x }
            },
            _ => {
                return Err(invalid_formula(text, "its operation is unrecognized"));
            },
        };
        Ok(formula)
    }
}

/// Convert the operand list into an exact-arity array, rejecting formulas
/// whose operand count does not match the operation.
fn take_operands<const COUNT: usize>(
    text: &str,
    operands: Vec<AdjustValue>,
) -> Result<[AdjustValue; COUNT]> {
    <[AdjustValue; COUNT]>::try_from(operands)
        .map_err(|_| invalid_formula(text, "it has the wrong operand count"))
}

fn invalid_formula(text: &str, reason: &str) -> Error {
    Error::Invalid(format!(
        "geometry guide formula '{text}' is invalid: {reason}"
    ))
}
