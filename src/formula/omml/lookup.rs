// Performance-optimized lookup tables for OMML parsing
//
// This module provides compile-time generated perfect hash function (PHF)
// lookup tables for fast element and attribute name resolution.

use super::elements::ElementType;
use phf::{phf_map, phf_set};
use crate::formula::ast::{Operator, AccentType, LargeOperator, PredefinedSymbol, FunctionName, Alignment, StyleType};

/// Fast element name to type lookup using PHF
pub static ELEMENT_TYPES: phf::Map<&'static str, ElementType> = phf_map! {
    // Core math elements
    "oMath" => ElementType::Math,
    "m:oMath" => ElementType::Math,

    // Text and runs
    "r" => ElementType::Run,
    "m:r" => ElementType::Run,
    "t" => ElementType::Text,
    "m:t" => ElementType::Text,

    // Functions
    "func" => ElementType::Function,
    "m:func" => ElementType::Function,
    "fName" => ElementType::FunctionName,
    "m:fName" => ElementType::FunctionName,

    // Fractions
    "f" => ElementType::Fraction,
    "m:f" => ElementType::Fraction,
    "num" => ElementType::Numerator,
    "m:num" => ElementType::Numerator,
    "den" => ElementType::Denominator,
    "m:den" => ElementType::Denominator,

    // Radicals
    "rad" => ElementType::Radical,
    "m:rad" => ElementType::Radical,
    "deg" => ElementType::Degree,
    "m:deg" => ElementType::Degree,

    // Scripts
    "sSup" => ElementType::Superscript,
    "m:sSup" => ElementType::Superscript,
    "sSub" => ElementType::Subscript,
    "m:sSub" => ElementType::Subscript,
    "sSubSup" => ElementType::SubSup,
    "m:sSubSup" => ElementType::SubSup,
    "sPre" => ElementType::PreScript,
    "m:sPre" => ElementType::PreScript,
    "sPost" => ElementType::PostScript,
    "m:sPost" => ElementType::PostScript,
    "sup" => ElementType::SuperscriptElement,
    "m:sup" => ElementType::SuperscriptElement,
    "sub" => ElementType::SubscriptElement,
    "m:sub" => ElementType::SubscriptElement,

    // Delimiters
    "d" => ElementType::Delimiter,
    "m:d" => ElementType::Delimiter,

    // N-ary operators
    "nary" => ElementType::Nary,
    "m:nary" => ElementType::Nary,
    "lim" => ElementType::Limit,
    "m:lim" => ElementType::Limit,
    "limLow" => ElementType::LimLow,
    "m:limLow" => ElementType::LimLow,
    "limUpp" => ElementType::LimUpp,
    "m:limUpp" => ElementType::LimUpp,

    // Matrices
    "m" => ElementType::Matrix,
    "m:m" => ElementType::Matrix,
    "mr" => ElementType::MatrixRow,
    "m:mr" => ElementType::MatrixRow,

    // Equation arrays
    "eqArr" => ElementType::EqArr,
    "m:eqArr" => ElementType::EqArr,

    // Accents and decorations
    "acc" => ElementType::Accent,
    "m:acc" => ElementType::Accent,
    "bar" => ElementType::Bar,
    "m:bar" => ElementType::Bar,
    "box" => ElementType::Box,
    "m:box" => ElementType::Box,
    "phant" => ElementType::Phantom,
    "m:phant" => ElementType::Phantom,
    "groupChr" => ElementType::GroupChar,
    "m:groupChr" => ElementType::GroupChar,
    "borderBox" => ElementType::BorderBox,
    "m:borderBox" => ElementType::BorderBox,

    // Properties and styles
    "rPr" => ElementType::Properties,
    "m:rPr" => ElementType::Properties,
    "fPr" => ElementType::Properties,
    "m:fPr" => ElementType::Properties,
    "dPr" => ElementType::Properties,
    "m:dPr" => ElementType::Properties,
    "naryPr" => ElementType::Properties,
    "m:naryPr" => ElementType::Properties,
    "accPr" => ElementType::AccentProperties,
    "m:accPr" => ElementType::AccentProperties,
    "radPr" => ElementType::Properties,
    "m:radPr" => ElementType::Properties,
    "sSupPr" => ElementType::Properties,
    "m:sSupPr" => ElementType::Properties,
    "sSubPr" => ElementType::Properties,
    "m:sSubPr" => ElementType::Properties,
    "funcPr" => ElementType::Properties,
    "m:funcPr" => ElementType::Properties,
    "groupChrPr" => ElementType::Properties,
    "m:groupChrPr" => ElementType::Properties,
    "eqArrPr" => ElementType::EqArrPr,
    "m:eqArrPr" => ElementType::EqArrPr,
    "boxPr" => ElementType::Properties,
    "m:boxPr" => ElementType::Properties,
    "phantPr" => ElementType::Properties,
    "m:phantPr" => ElementType::Properties,
    "borderBoxPr" => ElementType::Properties,
    "m:borderBoxPr" => ElementType::Properties,
    "mPr" => ElementType::Properties,
    "m:mPr" => ElementType::Properties,

    // Characters and content
    "chr" => ElementType::Character,
    "m:chr" => ElementType::Character,
    "begChr" => ElementType::Character,
    "m:begChr" => ElementType::Character,
    "endChr" => ElementType::Character,
    "m:endChr" => ElementType::Character,
    "e" => ElementType::Base,
    "m:e" => ElementType::Base,

    // Spacing
};

/// Fast operator character to operator lookup
pub static OPERATORS: phf::Map<&'static str, Operator> = phf_map! {
    // Basic operators
    "+" => Operator::Plus,
    "−" => Operator::Minus,
    "±" => Operator::PlusMinus,
    "∓" => Operator::MinusPlus,
    "×" => Operator::Multiply,
    "÷" => Operator::Divide,
    "⋅" => Operator::Dot,
    "∘" => Operator::Circle,
    "∗" => Operator::Star,
    "⋆" => Operator::Star,

    // Comparison operators
    "=" => Operator::Equals,
    "≠" => Operator::NotEquals,
    "<" => Operator::LessThan,
    ">" => Operator::GreaterThan,
    "≤" => Operator::LessThanOrEqual,
    "≥" => Operator::GreaterThanOrEqual,

    // Set operators
    "∈" => Operator::In,
    "∉" => Operator::NotIn,
    "⊂" => Operator::Subset,
    "⊃" => Operator::Superset,
    "⊆" => Operator::SubsetEq,
    "⊇" => Operator::SupersetEq,
    "∅" => Operator::EmptySet,
    "∪" => Operator::Union,
    "∩" => Operator::Intersection,

    // Relations
    "∼" => Operator::Approx,
    "≅" => Operator::Cong,
    "≡" => Operator::Equiv,
    "∝" => Operator::Propto,
    "∥" => Operator::Parallel,
    "⊥" => Operator::Perpendicular,

    // Arrows
    "←" => Operator::LeftArrow,
    "→" => Operator::RightArrow,
    "↑" => Operator::UpArrow,
    "↓" => Operator::DownArrow,
    "↔" => Operator::LeftRightArrow,
    "↕" => Operator::UpDownArrow,

    // Logical operators
    "∀" => Operator::ForAll,
    "∃" => Operator::Exists,
    "¬" => Operator::Not,
    "∧" => Operator::And,
    "∨" => Operator::Or,
    "⇒" => Operator::Implies,
    "⇔" => Operator::Iff,

    // Special symbols
    "∞" => Operator::Infinity,
    "ℵ" => Operator::Aleph,
    "′" => Operator::Prime,
    "″" => Operator::DoublePrime,
    "‴" => Operator::TriplePrime,
    "…" => Operator::Ellipsis,
    "⋯" => Operator::CDots,
    "⋮" => Operator::VDots,
    "⋱" => Operator::DDots,

    // Calculus
    "∇" => Operator::Nabla,
    "∂" => Operator::Partial,
    "∫" => Operator::Differential,

    // Miscellaneous
    "∴" => Operator::Therefore,
    "∵" => Operator::Because,
    "□" => Operator::Box,
    "◆" => Operator::Diamond,
    "■" => Operator::Square,
};

/// Fast accent character to accent type lookup
pub static ACCENTS: phf::Map<&'static str, AccentType> = phf_map! {
    "¯" => AccentType::Bar,
    "¨" => AccentType::DoubleDot,
    "˙" => AccentType::Dot,
    "`" => AccentType::Grave,
    "´" => AccentType::Acute,
    "˜" => AccentType::Tilde,
    "^" => AccentType::Hat,
    "ˇ" => AccentType::Check,
    "˘" => AccentType::Breve,
    "→" => AccentType::Vec,
    "⃛" => AccentType::TripleDot,
};

/// Fast large operator character to operator lookup
pub static LARGE_OPERATORS: phf::Map<&'static str, LargeOperator> = phf_map! {
    "∑" => LargeOperator::Sum,
    "∏" => LargeOperator::Product,
    "∐" => LargeOperator::Coproduct,
    "∫" => LargeOperator::Integral,
    "∬" => LargeOperator::DoubleIntegral,
    "∭" => LargeOperator::TripleIntegral,
    "∮" => LargeOperator::ContourIntegral,
    "∯" => LargeOperator::SurfaceIntegral,
    "∰" => LargeOperator::VolumeIntegral,
    "⋃" => LargeOperator::Union,
    "⋂" => LargeOperator::Intersection,
    "⨄" => LargeOperator::BigUnion,
    "⨅" => LargeOperator::BigIntersection,
    "lim" => LargeOperator::Limit,
    "max" => LargeOperator::Max,
    "min" => LargeOperator::Min,
    "sup" => LargeOperator::Supremum,
    "inf" => LargeOperator::Infimum,
    "argmax" => LargeOperator::ArgMax,
    "argmin" => LargeOperator::ArgMin,
};

/// Fast predefined symbol lookup
pub static PREDEFINED_SYMBOLS: phf::Map<&'static str, PredefinedSymbol> = phf_map! {
    // Greek lowercase
    "α" => PredefinedSymbol::Alpha,
    "β" => PredefinedSymbol::Beta,
    "γ" => PredefinedSymbol::Gamma,
    "δ" => PredefinedSymbol::Delta,
    "ε" => PredefinedSymbol::Epsilon,
    "ζ" => PredefinedSymbol::Zeta,
    "η" => PredefinedSymbol::Eta,
    "θ" => PredefinedSymbol::Theta,
    "ι" => PredefinedSymbol::Iota,
    "κ" => PredefinedSymbol::Kappa,
    "λ" => PredefinedSymbol::Lambda,
    "μ" => PredefinedSymbol::Mu,
    "ν" => PredefinedSymbol::Nu,
    "ξ" => PredefinedSymbol::Xi,
    "ο" => PredefinedSymbol::Omicron,
    "π" => PredefinedSymbol::Pi,
    "ρ" => PredefinedSymbol::Rho,
    "σ" => PredefinedSymbol::Sigma,
    "τ" => PredefinedSymbol::Tau,
    "υ" => PredefinedSymbol::Upsilon,
    "φ" => PredefinedSymbol::Phi,
    "χ" => PredefinedSymbol::Chi,
    "ψ" => PredefinedSymbol::Psi,
    "ω" => PredefinedSymbol::Omega,

    // Greek uppercase
    "Α" => PredefinedSymbol::AlphaCap,
    "Β" => PredefinedSymbol::BetaCap,
    "Γ" => PredefinedSymbol::GammaCap,
    "Δ" => PredefinedSymbol::DeltaCap,
    "Ε" => PredefinedSymbol::EpsilonCap,
    "Ζ" => PredefinedSymbol::ZetaCap,
    "Η" => PredefinedSymbol::EtaCap,
    "Θ" => PredefinedSymbol::ThetaCap,
    "Ι" => PredefinedSymbol::IotaCap,
    "Κ" => PredefinedSymbol::KappaCap,
    "Λ" => PredefinedSymbol::LambdaCap,
    "Μ" => PredefinedSymbol::MuCap,
    "Ν" => PredefinedSymbol::NuCap,
    "Ξ" => PredefinedSymbol::XiCap,
    "Ο" => PredefinedSymbol::OmicronCap,
    "Π" => PredefinedSymbol::PiCap,
    "Ρ" => PredefinedSymbol::RhoCap,
    "Σ" => PredefinedSymbol::SigmaCap,
    "Τ" => PredefinedSymbol::TauCap,
    "Υ" => PredefinedSymbol::UpsilonCap,
    "Φ" => PredefinedSymbol::PhiCap,
    "Χ" => PredefinedSymbol::ChiCap,
    "Ψ" => PredefinedSymbol::PsiCap,
    "Ω" => PredefinedSymbol::OmegaCap,

    // Hebrew
    "ℵ" => PredefinedSymbol::Aleph,

    // Constants
    "e" => PredefinedSymbol::ExponentialE,
    "i" => PredefinedSymbol::ImaginaryI,

    // Infinity
    "∞" => PredefinedSymbol::Infinity,
};

/// Fast function name lookup
pub static FUNCTIONS: phf::Map<&'static str, FunctionName> = phf_map! {
    "sin" => FunctionName::Sin,
    "cos" => FunctionName::Cos,
    "tan" => FunctionName::Tan,
    "sec" => FunctionName::Sec,
    "csc" => FunctionName::Csc,
    "cot" => FunctionName::Cot,
    "arcsin" => FunctionName::ArcSin,
    "arccos" => FunctionName::ArcCos,
    "arctan" => FunctionName::ArcTan,
    "arcsec" => FunctionName::ArcSec,
    "arccsc" => FunctionName::ArcCsc,
    "arccot" => FunctionName::ArcCot,
    "sinh" => FunctionName::Sinh,
    "cosh" => FunctionName::Cosh,
    "tanh" => FunctionName::Tanh,
    "sech" => FunctionName::Sech,
    "csch" => FunctionName::Csch,
    "coth" => FunctionName::Coth,
    "log" => FunctionName::Log,
    "ln" => FunctionName::Ln,
    "exp" => FunctionName::Exp,
    "sqrt" => FunctionName::Sqrt,
    "min" => FunctionName::Min,
    "max" => FunctionName::Max,
    "sup" => FunctionName::Sup,
    "inf" => FunctionName::Inf,
    "lim" => FunctionName::Lim,
    "det" => FunctionName::Det,
    "trace" => FunctionName::Trace,
    "dim" => FunctionName::Dim,
    "ker" => FunctionName::Ker,
    "Im" => FunctionName::Im,
    "Re" => FunctionName::Re,
    "arg" => FunctionName::Arg,
    "mod" => FunctionName::Mod,
    "gcd" => FunctionName::Gcd,
    "lcm" => FunctionName::Lcm,
};

/// Fast attribute value parsing lookup tables
pub static BOOLEAN_VALUES: phf::Set<&'static str> = phf_set! {
    "1", "true", "on", "yes"
};

pub static ALIGNMENT_VALUES: phf::Map<&'static str, Alignment> = phf_map! {
    "left" => Alignment::Left,
    "center" => Alignment::Center,
    "right" => Alignment::Right,
    "top" => Alignment::Top,
    "bottom" => Alignment::Bottom,
    "baseline" => Alignment::Baseline,
    "axis" => Alignment::Axis,
    "centered" => Alignment::Centered,
    "match" => Alignment::Match,
};

pub static STYLE_VALUES: phf::Map<&'static str, StyleType> = phf_map! {
    "p" => StyleType::Normal,
    "roman" => StyleType::Normal,
    "b" => StyleType::Bold,
    "bold" => StyleType::Bold,
    "i" => StyleType::Italic,
    "italic" => StyleType::Italic,
    "bi" => StyleType::BoldItalic,
    "bold-italic" => StyleType::BoldItalic,
    "ss" => StyleType::SansSerif,
    "sans-serif" => StyleType::SansSerif,
    "ssb" => StyleType::SansSerifBold,
    "sans-serif-bold" => StyleType::SansSerifBold,
    "ssi" => StyleType::SansSerifItalic,
    "sans-serif-italic" => StyleType::SansSerifItalic,
    "ssbi" => StyleType::SansSerifBoldItalic,
    "sans-serif-bold-italic" => StyleType::SansSerifBoldItalic,
    "m" => StyleType::Monospace,
    "monospace" => StyleType::Monospace,
    "sc" => StyleType::Script,
    "script" => StyleType::Script,
    "bsc" => StyleType::BoldScript,
    "bold-script" => StyleType::BoldScript,
    "fr" => StyleType::Fraktur,
    "fraktur" => StyleType::Fraktur,
    "bfr" => StyleType::BoldFraktur,
    "bold-fraktur" => StyleType::BoldFraktur,
    "ds" => StyleType::DoubleStruck,
    "double-struck" => StyleType::DoubleStruck,
};

/// Fast lookup functions using PHF maps
/// Get element type from name with fallback to Unknown
pub fn get_element_type(name: &str) -> ElementType {
    ELEMENT_TYPES.get(name).copied().unwrap_or(ElementType::Unknown)
}

/// Get operator from character
pub fn get_operator(chr: &str) -> Option<Operator> {
    OPERATORS.get(chr).copied()
}

/// Get accent type from character
pub fn get_accent_type(chr: &str) -> Option<AccentType> {
    ACCENTS.get(chr).copied()
}

/// Get large operator from character
pub fn get_large_operator(chr: &str) -> Option<LargeOperator> {
    LARGE_OPERATORS.get(chr).copied()
}

/// Get predefined symbol from character
pub fn get_predefined_symbol(chr: &str) -> Option<PredefinedSymbol> {
    PREDEFINED_SYMBOLS.get(chr).copied()
}

/// Get function name from string
pub fn get_function_name(name: &str) -> Option<FunctionName> {
    FUNCTIONS.get(name).copied()
}

/// Parse boolean attribute value
pub fn parse_bool_value(val: &str) -> Option<bool> {
    if BOOLEAN_VALUES.contains(val) {
        Some(true)
    } else if val == "0" || val == "false" || val == "off" || val == "no" {
        Some(false)
    } else {
        None
    }
}

/// Parse alignment value
pub fn parse_alignment_value(val: &str) -> Option<Alignment> {
    ALIGNMENT_VALUES.get(val).copied()
}

/// Parse style value
pub fn parse_style_value(val: &str) -> Option<StyleType> {
    STYLE_VALUES.get(val).copied()
}
