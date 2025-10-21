use std::borrow::Cow;

/// Mathematical operators
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Operator {
    // Basic arithmetic
    Plus,
    Minus,
    Multiply,
    Divide,
    PlusMinus,
    MinusPlus,

    // Comparison
    Equals,
    NotEquals,
    LessThan,
    GreaterThan,
    LessThanOrEqual,
    GreaterThanOrEqual,

    // Algebraic
    Times,
    Dot,
    Cross,
    Star,
    Circle,
    Circ,
    Bullet,
    Wedge,
    Vee,
    Cap,
    Cup,

    // Set theory
    In,
    NotIn,
    Subset,
    Superset,
    SubsetEq,
    SupersetEq,
    EmptySet,
    Union,
    Intersection,

    // Relations
    Approx,
    Cong,
    Equiv,
    Propto,
    Sim,
    Simeq,
    Asymp,

    // Geometry
    Parallel,
    Perpendicular,
    Angle,

    // Calculus
    Nabla,
    Partial,
    Differential,

    // Special symbols
    Infinity,
    Aleph,
    Prime,
    DoublePrime,
    TriplePrime,

    // Dots
    Ellipsis,
    CDots,
    VDots,
    DDots,
    Ldots,

    // Arrows
    LeftArrow,
    RightArrow,
    UpArrow,
    DownArrow,
    LeftRightArrow,
    UpDownArrow,

    // Logical
    ForAll,
    Exists,
    Not,
    And,
    Or,
    Implies,
    Iff,

    // Miscellaneous
    Therefore,
    Because,
    Box,
    Diamond,
    Square,
}

/// Mathematical symbols (Greek letters, special symbols, etc.)
#[derive(Debug, Clone, PartialEq)]
pub struct Symbol<'a> {
    pub name: Cow<'a, str>,
    pub unicode: Option<char>,
    pub variant: Option<StyleType>,
}

/// Function names for mathematical functions
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FunctionName {
    Sin,
    Cos,
    Tan,
    Sec,
    Csc,
    Cot,
    ArcSin,
    ArcCos,
    ArcTan,
    ArcSec,
    ArcCsc,
    ArcCot,
    Sinh,
    Cosh,
    Tanh,
    Sech,
    Csch,
    Coth,
    Log,
    Ln,
    Exp,
    Sqrt,
    Min,
    Max,
    Sup,
    Inf,
    Lim,
    Det,
    Trace,
    Dim,
    Ker,
    Im,
    Re,
    Arg,
    Mod,
    Gcd,
    Lcm,
}

/// Predefined symbols for common mathematical entities
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PredefinedSymbol {
    // Greek letters
    Alpha,
    Beta,
    Gamma,
    Delta,
    Epsilon,
    Zeta,
    Eta,
    Theta,
    Iota,
    Kappa,
    Lambda,
    Mu,
    Nu,
    Xi,
    Omicron,
    Pi,
    Rho,
    Sigma,
    Tau,
    Upsilon,
    Phi,
    Chi,
    Psi,
    Omega,

    // Capital Greek
    AlphaCap,
    BetaCap,
    GammaCap,
    DeltaCap,
    EpsilonCap,
    ZetaCap,
    EtaCap,
    ThetaCap,
    IotaCap,
    KappaCap,
    LambdaCap,
    MuCap,
    NuCap,
    XiCap,
    OmicronCap,
    PiCap,
    RhoCap,
    SigmaCap,
    TauCap,
    UpsilonCap,
    PhiCap,
    ChiCap,
    PsiCap,
    OmegaCap,

    // Hebrew
    Aleph,

    // Special constants
    EulerGamma,
    ExponentialE,
    ImaginaryI,

    // Infinity
    Infinity,
}

/// Fence types for fenced expressions
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Fence {
    Paren,         // ( )
    Bracket,       // [ ]
    Brace,         // { }
    Angle,         // ⟨ ⟩
    Pipe,          // | |
    DoublePipe,    // ‖ ‖
    Floor,         // ⌊ ⌋
    Ceiling,       // ⌈ ⌉
    AngleBracket,  // 〈 〉
    SquareBracket, // ⟦ ⟧
    CurlyBrace,    // ⦃ ⦄
    None,          // No fence
}

/// Fence character specification for customizable fences
#[derive(Debug, Clone, PartialEq)]
pub struct FenceSpec {
    pub open: Option<String>,
    pub close: Option<String>,
}

/// Large operators (sum, product, integral, etc.)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LargeOperator {
    Sum,             // ∑
    Product,         // ∏
    Coproduct,       // ∐
    Integral,        // ∫
    DoubleIntegral,  // ∬
    TripleIntegral,  // ∭
    ContourIntegral, // ∮
    SurfaceIntegral, // ∯
    VolumeIntegral,  // ∰
    Union,           // ⋃
    Intersection,    // ⋂
    BigUnion,        // ⋃
    BigIntersection, // ⋂
    Limit,           // lim
    Max,             // max
    Min,             // min
    Supremum,        // sup
    Infimum,         // inf
    ArgMax,          // argmax
    ArgMin,          // argmin
}

/// Matrix fence types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MatrixFence {
    None,       // No fence
    Paren,      // ( )
    Bracket,    // [ ]
    Brace,      // { }
    Pipe,       // | |
    DoublePipe, // ‖ ‖
}

/// Accent types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AccentType {
    Hat,       // ^
    Check,     // ˇ
    Tilde,     // ~
    Acute,     // ´
    Grave,     // `
    Dot,       // ˙
    DoubleDot, // ¨
    TripleDot, // ⃛
    Bar,       // ¯
    Breve,     // ˘
    Vec,       // →
}

/// Space types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SpaceType {
    Thin,     // Thin space
    Medium,   // Medium space
    Thick,    // Thick space
    Quad,     // Quad space
    QQuad,    // Double quad space
    Negative, // Negative space
}

/// Style types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleType {
    Normal,
    Bold,
    Italic,
    BoldItalic,
    SansSerif,
    SansSerifBold,
    SansSerifItalic,
    SansSerifBoldItalic,
    Monospace,
    Script,
    BoldScript,
    Fraktur,
    BoldFraktur,
    DoubleStruck,
}

/// Alignment types for positioning elements
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Top,
    Bottom,
    Baseline,
    Axis,
    Centered,
    Match,
}

/// Vertical alignment types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
    Top,
    Bottom,
    Center,
    Baseline,
    Axis,
}

/// Position types for scripts and accents
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Position {
    Prefix,
    Postfix,
    Infix,
    Top,
    Bottom,
}

/// Fraction types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FractionType {
    Bar,    // Normal fraction bar
    NoBar,  // Linear fraction (no bar)
    Skewed, // Skewed fraction bar
}

/// Math variant types (equivalent to style types but OMML-specific)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathVariant {
    Normal,
    Bold,
    Italic,
    BoldItalic,
    SansSerif,
    SansSerifBold,
    SansSerifItalic,
    SansSerifBoldItalic,
    Monospace,
    Script,
    BoldScript,
    Fraktur,
    BoldFraktur,
    DoubleStruck,
}

/// Shape types for delimiters
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeType {
    Centered,
    Match,
}

/// Break types for line breaking
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BreakType {
    None,
    Line,
    Page,
}

/// Underline/overline styles
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineStyle {
    Single,
    Double,
    Thick,
    Dotted,
    Dashed,
    Wave,
}

/// Strike-through styles
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StrikeStyle {
    Single,
    Double,
}

/// Border box properties
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BorderBoxStyle {
    pub hide_top: bool,
    pub hide_bottom: bool,
    pub hide_left: bool,
    pub hide_right: bool,
    pub strike_horizontal: bool,
    pub strike_vertical: bool,
    pub strike_bltr: bool, // bottom-left to top-right
    pub strike_tlbr: bool, // top-left to bottom-right
}

/// Equation array properties
#[derive(Debug, Clone, PartialEq)]
pub struct EqArrayProperties {
    pub base_alignment: Option<Alignment>,
    pub max_distance: Option<f32>,
    pub object_distance: Option<f32>,
    pub row_spacing: Option<f32>,
    pub row_spacing_rule: Option<String>,
}

/// Matrix properties
#[derive(Debug, Clone, PartialEq)]
pub struct MatrixProperties {
    pub base_alignment: Option<Alignment>,
    pub column_gap: Option<f32>,
    pub row_spacing: Option<f32>,
    pub column_spacing: Option<Vec<f32>>,
}

/// Types of limits in mathematical expressions
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LimitType {
    Lower,
    Upper,
}
