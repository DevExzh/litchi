// Math node definitions

use super::types::*;
use std::borrow::Cow;

/// Math node representing a single element in the formula AST
#[derive(Debug, Clone, PartialEq)]
pub enum MathNode<'a> {
    /// Plain text or identifier
    Text(Cow<'a, str>),

    /// Numeric value
    Number(Cow<'a, str>),

    /// Mathematical operator
    Operator(Operator),

    /// Symbol (Greek letters, special symbols, etc.)
    Symbol(Symbol<'a>),

    /// Predefined symbol (Greek letters, constants, etc.)
    PredefinedSymbol(PredefinedSymbol),

    /// Fraction: numerator / denominator
    Frac {
        numerator: Vec<MathNode<'a>>,
        denominator: Vec<MathNode<'a>>,
        line_thickness: Option<f32>,
        frac_type: Option<FractionType>,
    },

    /// Square root or nth root
    Root {
        base: Vec<MathNode<'a>>,
        index: Option<Vec<MathNode<'a>>>,
    },

    /// Superscript (power)
    Power {
        base: Vec<MathNode<'a>>,
        exponent: Vec<MathNode<'a>>,
    },

    /// Subscript
    Sub {
        base: Vec<MathNode<'a>>,
        subscript: Vec<MathNode<'a>>,
    },

    /// Both subscript and superscript
    SubSup {
        base: Vec<MathNode<'a>>,
        subscript: Vec<MathNode<'a>>,
        superscript: Vec<MathNode<'a>>,
    },

    /// Pre-subscript
    PreSub {
        base: Vec<MathNode<'a>>,
        pre_subscript: Vec<MathNode<'a>>,
    },

    /// Pre-superscript
    PreSup {
        base: Vec<MathNode<'a>>,
        pre_superscript: Vec<MathNode<'a>>,
    },

    /// Pre-subscript and pre-superscript
    PreSubSup {
        base: Vec<MathNode<'a>>,
        pre_subscript: Vec<MathNode<'a>>,
        pre_superscript: Vec<MathNode<'a>>,
    },

    /// Underscript
    Under {
        base: Vec<MathNode<'a>>,
        under: Vec<MathNode<'a>>,
        position: Option<Position>,
    },

    /// Overscript
    Over {
        base: Vec<MathNode<'a>>,
        over: Vec<MathNode<'a>>,
        position: Option<Position>,
    },

    /// Both underscript and overscript
    UnderOver {
        base: Vec<MathNode<'a>>,
        under: Vec<MathNode<'a>>,
        over: Vec<MathNode<'a>>,
        position: Option<Position>,
    },

    /// Fenced expression (parentheses, brackets, etc.)
    Fenced {
        open: Fence,
        content: Vec<MathNode<'a>>,
        close: Fence,
        separator: Option<Cow<'a, str>>,
    },

    /// Function with limits (sum, product, integral, etc.)
    LargeOp {
        operator: LargeOperator,
        lower_limit: Option<Vec<MathNode<'a>>>,
        upper_limit: Option<Vec<MathNode<'a>>>,
        integrand: Option<Vec<MathNode<'a>>>,
        hide_lower: bool,
        hide_upper: bool,
    },

    /// Named function (sin, cos, log, etc.)
    Function {
        name: Cow<'a, str>,
        argument: Vec<MathNode<'a>>,
    },

    /// Predefined function
    PredefinedFunction {
        function: FunctionName,
        argument: Vec<MathNode<'a>>,
    },

    /// Matrix or array
    Matrix {
        rows: Vec<Vec<Vec<MathNode<'a>>>>,
        fence_type: MatrixFence,
        properties: Option<MatrixProperties>,
    },

    /// Equation array (aligned equations)
    EqArray {
        rows: Vec<Vec<MathNode<'a>>>,
        properties: Option<EqArrayProperties>,
    },

    /// Accent over base
    Accent {
        base: Box<Vec<MathNode<'a>>>,
        accent: AccentType,
        position: Option<Position>,
    },

    /// Bar (overline)
    Bar {
        base: Box<Vec<MathNode<'a>>>,
        position: Option<Position>,
    },

    /// Box (border box)
    BorderBox {
        content: Box<Vec<MathNode<'a>>>,
        style: Option<BorderBoxStyle>,
    },

    /// Group character (brace, bracket over/under)
    GroupChar {
        base: Box<Vec<MathNode<'a>>>,
        character: Option<Cow<'a, str>>,
        position: Option<Position>,
        vertical_alignment: Option<VerticalAlignment>,
    },

    /// Space
    Space(SpaceType),

    /// Line break
    LineBreak,

    /// Style change
    Style {
        style: StyleType,
        content: Vec<MathNode<'a>>,
    },

    /// Run with properties
    Run {
        content: Vec<MathNode<'a>>,
        literal: Option<bool>,
        style: Option<StyleType>,
        font: Option<Cow<'a, str>>,
        color: Option<Cow<'a, str>>,
        underline: Option<LineStyle>,
        overline: Option<LineStyle>,
        strike_through: Option<StrikeStyle>,
        double_strike_through: Option<bool>,
    },

    /// Row of elements
    Row(Vec<MathNode<'a>>),

    /// Phantom (invisible content that takes up space)
    Phantom(Box<Vec<MathNode<'a>>>),

    /// Limit element (lower/upper limit for operators)
    Limit {
        content: Box<Vec<MathNode<'a>>>,
        limit_type: LimitType,
    },

    /// Degree element (for radicals)
    Degree(Box<Vec<MathNode<'a>>>),

    /// Base element (for scripts and operators)
    Base(Box<Vec<MathNode<'a>>>),

    /// Argument element (for functions)
    Argument(Box<Vec<MathNode<'a>>>),

    /// Numerator element
    Numerator(Box<Vec<MathNode<'a>>>),

    /// Denominator element
    Denominator(Box<Vec<MathNode<'a>>>),

    /// Integrand element (for integrals)
    Integrand(Box<Vec<MathNode<'a>>>),

    /// Lower limit element
    LowerLimit(Box<Vec<MathNode<'a>>>),

    /// Upper limit element
    UpperLimit(Box<Vec<MathNode<'a>>>),

    /// Error node for invalid formulas
    Error(Cow<'a, str>),
}


