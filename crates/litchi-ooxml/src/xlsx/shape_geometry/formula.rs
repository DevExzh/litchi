//! Typed DrawingML geometry-guide formulas (`a:gd@fmla`,
//! ST_GeomGuideFormula; ECMA-376 part 1 §20.1.10.11).
//!
//! A formula is an operation token followed by space-delimited operands,
//! each a literal or a guide reference. [`XlsxGeometryFormula`] models every
//! defined operation with its exact operand count; parsing rejects unknown
//! operations and wrong arities so authored formulas always serialize back
//! to a schema-valid token.

use std::fmt;
use std::str::FromStr;

use crate::error::{OoxmlError, Result};

use super::XlsxAdjustValue;

/// A geometry-guide formula: one ECMA-376 operation with typed operands.
///
/// Formula semantics quote §20.1.10.11; trigonometric operands are angles in
/// 60000ths of a degree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsxGeometryFormula {
    /// `*/ x y z` — `(x * y) / z`.
    MultiplyDivide {
        /// First factor.
        x: XlsxAdjustValue,
        /// Second factor.
        y: XlsxAdjustValue,
        /// Divisor.
        z: XlsxAdjustValue,
    },
    /// `+- x y z` — `(x + y) - z`.
    AddSubtract {
        /// First addend.
        x: XlsxAdjustValue,
        /// Second addend.
        y: XlsxAdjustValue,
        /// Subtrahend.
        z: XlsxAdjustValue,
    },
    /// `+/ x y z` — `(x + y) / z`.
    AddDivide {
        /// First addend.
        x: XlsxAdjustValue,
        /// Second addend.
        y: XlsxAdjustValue,
        /// Divisor.
        z: XlsxAdjustValue,
    },
    /// `?: x y z` — `y` when `x > 0`, `z` otherwise.
    IfElse {
        /// Condition value.
        x: XlsxAdjustValue,
        /// Result when the condition is positive.
        y: XlsxAdjustValue,
        /// Result otherwise.
        z: XlsxAdjustValue,
    },
    /// `abs x` — `|x|`.
    Absolute {
        /// Input value.
        x: XlsxAdjustValue,
    },
    /// `at2 x y` — `atan2(y, x)`.
    ArcTangent {
        /// Horizontal component.
        x: XlsxAdjustValue,
        /// Vertical component.
        y: XlsxAdjustValue,
    },
    /// `cat2 x y z` — `x * cos(atan2(z, y))`.
    CosineArcTangent {
        /// Scale value.
        x: XlsxAdjustValue,
        /// Horizontal component.
        y: XlsxAdjustValue,
        /// Vertical component.
        z: XlsxAdjustValue,
    },
    /// `cos x y` — `x * cos(y)`.
    Cosine {
        /// Scale value.
        x: XlsxAdjustValue,
        /// Angle.
        y: XlsxAdjustValue,
    },
    /// `max x y` — the greater of `x` and `y`.
    Maximum {
        /// First value.
        x: XlsxAdjustValue,
        /// Second value.
        y: XlsxAdjustValue,
    },
    /// `min x y` — the lesser of `x` and `y`.
    Minimum {
        /// First value.
        x: XlsxAdjustValue,
        /// Second value.
        y: XlsxAdjustValue,
    },
    /// `mod x y z` — `sqrt(x² + y² + z²)`.
    Modulus {
        /// First component.
        x: XlsxAdjustValue,
        /// Second component.
        y: XlsxAdjustValue,
        /// Third component.
        z: XlsxAdjustValue,
    },
    /// `pin x y z` — `y` clamped to `[x, z]`.
    Pin {
        /// Lower bound.
        x: XlsxAdjustValue,
        /// Clamped value.
        y: XlsxAdjustValue,
        /// Upper bound.
        z: XlsxAdjustValue,
    },
    /// `sat2 x y z` — `x * sin(atan2(z, y))`.
    SineArcTangent {
        /// Scale value.
        x: XlsxAdjustValue,
        /// Horizontal component.
        y: XlsxAdjustValue,
        /// Vertical component.
        z: XlsxAdjustValue,
    },
    /// `sin x y` — `x * sin(y)`.
    Sine {
        /// Scale value.
        x: XlsxAdjustValue,
        /// Angle.
        y: XlsxAdjustValue,
    },
    /// `sqrt x` — `sqrt(x)`.
    SquareRoot {
        /// Input value.
        x: XlsxAdjustValue,
    },
    /// `tan x y` — `x * tan(y)`.
    Tangent {
        /// Scale value.
        x: XlsxAdjustValue,
        /// Angle.
        y: XlsxAdjustValue,
    },
    /// `val x` — the literal or referenced value itself.
    Value {
        /// The value.
        x: XlsxAdjustValue,
    },
}

impl XlsxGeometryFormula {
    /// A `val` formula holding one literal, the common shape of adjust-value
    /// entries.
    pub fn literal(value: i64) -> Self {
        Self::Value {
            x: XlsxAdjustValue::Value(value),
        }
    }

    /// The formula's operands in serialization order.
    pub fn operands(&self) -> impl Iterator<Item = &XlsxAdjustValue> {
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

    /// The formula's ST_GeomGuideFormula operation token.
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
    operands: [Option<&'formula XlsxAdjustValue>; 3],
    next: usize,
}

impl<'formula> FormulaOperands<'formula> {
    fn one(x: &'formula XlsxAdjustValue) -> Self {
        Self {
            operands: [Some(x), None, None],
            next: 0,
        }
    }

    fn two(x: &'formula XlsxAdjustValue, y: &'formula XlsxAdjustValue) -> Self {
        Self {
            operands: [Some(x), Some(y), None],
            next: 0,
        }
    }

    fn three(
        x: &'formula XlsxAdjustValue,
        y: &'formula XlsxAdjustValue,
        z: &'formula XlsxAdjustValue,
    ) -> Self {
        Self {
            operands: [Some(x), Some(y), Some(z)],
            next: 0,
        }
    }
}

impl<'formula> Iterator for FormulaOperands<'formula> {
    type Item = &'formula XlsxAdjustValue;

    fn next(&mut self) -> Option<Self::Item> {
        let operand = self.operands.get(self.next).copied().flatten()?;
        self.next += 1;
        Some(operand)
    }
}

impl fmt::Display for XlsxGeometryFormula {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.operation())?;
        for operand in self.operands() {
            write!(formatter, " {operand}")?;
        }
        Ok(())
    }
}

impl FromStr for XlsxGeometryFormula {
    type Err = OoxmlError;

    fn from_str(text: &str) -> Result<Self> {
        let mut tokens = text.split_whitespace();
        let operation = tokens
            .next()
            .ok_or_else(|| invalid_formula(text, "it has no operation"))?;
        let operands = tokens
            .map(XlsxAdjustValue::from_str)
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
    operands: Vec<XlsxAdjustValue>,
) -> Result<[XlsxAdjustValue; COUNT]> {
    <[XlsxAdjustValue; COUNT]>::try_from(operands)
        .map_err(|_| invalid_formula(text, "it has the wrong operand count"))
}

fn invalid_formula(text: &str, reason: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(format!(
        "geometry guide formula '{text}' is invalid: {reason}"
    ))
}
