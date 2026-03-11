// Operator and symbol conversion to LaTeX

use crate::formula::ast::{AccentType, Fence, LargeOperator, Operator, SpaceType, StyleType};

/// Convert operator to LaTeX string
pub fn operator_to_latex(op: Operator) -> &'static str {
    match op {
        Operator::Plus => "+",
        Operator::Minus => "-",
        Operator::Multiply => "\\cdot",
        Operator::Divide => "\\div",
        Operator::Equals => "=",
        Operator::NotEquals => "\\neq",
        Operator::LessThan => "<",
        Operator::GreaterThan => ">",
        Operator::LessThanOrEqual => "\\leq",
        Operator::GreaterThanOrEqual => "\\geq",
        Operator::PlusMinus => "\\pm",
        Operator::MinusPlus => "\\mp",
        Operator::Times => "\\times",
        Operator::Dot => "\\cdot",
        Operator::Cross => "\\times",
        Operator::Star => "\\ast",
        Operator::Circle => "\\circ",
        Operator::Circ => "\\circ",
        Operator::Bullet => "\\bullet",
        Operator::Wedge => "\\wedge",
        Operator::Vee => "\\vee",
        Operator::Cap => "\\cap",
        Operator::Cup => "\\cup",
        Operator::In => "\\in",
        Operator::NotIn => "\\notin",
        Operator::Subset => "\\subset",
        Operator::Superset => "\\supset",
        Operator::SubsetEq => "\\subseteq",
        Operator::SupersetEq => "\\supseteq",
        Operator::Approx => "\\approx",
        Operator::Cong => "\\cong",
        Operator::Equiv => "\\equiv",
        Operator::Propto => "\\propto",
        Operator::Parallel => "\\parallel",
        Operator::Perpendicular => "\\perp",
        Operator::Angle => "\\angle",
        Operator::Nabla => "\\nabla",
        Operator::Partial => "\\partial",
        Operator::Infinity => "\\infty",
        Operator::Aleph => "\\aleph",
        Operator::Prime => "'",
        Operator::DoublePrime => "''",
        Operator::Ellipsis => "\\ldots",
        Operator::CDots => "\\cdots",
        Operator::VDots => "\\vdots",
        Operator::DDots => "\\ddots",
        Operator::Ldots => "\\ldots",
        Operator::EmptySet => "\\emptyset",
        Operator::Union => "\\cup",
        Operator::Intersection => "\\cap",
        Operator::Sim => "\\sim",
        Operator::Simeq => "\\simeq",
        Operator::Asymp => "\\asymp",
        Operator::Differential => "\\mathrm{d}",
        Operator::TriplePrime => "'''",
        Operator::LeftArrow => "\\leftarrow",
        Operator::RightArrow => "\\rightarrow",
        Operator::UpArrow => "\\uparrow",
        Operator::DownArrow => "\\downarrow",
        Operator::LeftRightArrow => "\\leftrightarrow",
        Operator::UpDownArrow => "\\updownarrow",
        Operator::ForAll => "\\forall",
        Operator::Exists => "\\exists",
        Operator::Not => "\\neg",
        Operator::And => "\\land",
        Operator::Or => "\\lor",
        Operator::Implies => "\\implies",
        Operator::Iff => "\\iff",
        Operator::Therefore => "\\therefore",
        Operator::Because => "\\because",
        Operator::Box => "\\Box",
        Operator::Diamond => "\\Diamond",
        Operator::Square => "\\square",
    }
}

/// Convert fence to LaTeX string
pub fn fence_to_latex(fence: Fence, is_open: bool) -> &'static str {
    match (fence, is_open) {
        (Fence::Paren, true) => "\\left(",
        (Fence::Paren, false) => "\\right)",
        (Fence::Bracket, true) => "\\left[",
        (Fence::Bracket, false) => "\\right]",
        (Fence::Brace, true) => "\\left\\{",
        (Fence::Brace, false) => "\\right\\}",
        (Fence::Angle, true) => "\\left\\langle",
        (Fence::Angle, false) => "\\right\\rangle",
        (Fence::Pipe, true) => "\\left|",
        (Fence::Pipe, false) => "\\right|",
        (Fence::DoublePipe, true) => "\\left\\|",
        (Fence::DoublePipe, false) => "\\right\\|",
        (Fence::Floor, true) => "\\left\\lfloor",
        (Fence::Floor, false) => "\\right\\rfloor",
        (Fence::Ceiling, true) => "\\left\\lceil",
        (Fence::Ceiling, false) => "\\right\\rceil",
        (Fence::AngleBracket, true) => "\\left\\langle",
        (Fence::AngleBracket, false) => "\\right\\rangle",
        (Fence::SquareBracket, true) => "\\left\\lbrack",
        (Fence::SquareBracket, false) => "\\right\\rbrack",
        (Fence::CurlyBrace, true) => "\\left\\lbrace",
        (Fence::CurlyBrace, false) => "\\right\\rbrace",
        (Fence::None, true) => "\\left.",
        (Fence::None, false) => "\\right.",
    }
}

/// Convert large operator to LaTeX string
pub fn large_operator_to_latex(op: LargeOperator) -> &'static str {
    match op {
        LargeOperator::Sum => "\\sum",
        LargeOperator::Product => "\\prod",
        LargeOperator::Coproduct => "\\coprod",
        LargeOperator::Integral => "\\int",
        LargeOperator::DoubleIntegral => "\\iint",
        LargeOperator::TripleIntegral => "\\iiint",
        LargeOperator::ContourIntegral => "\\oint",
        LargeOperator::SurfaceIntegral => "\\oiint",
        LargeOperator::VolumeIntegral => "\\oiiint",
        LargeOperator::Union => "\\bigcup",
        LargeOperator::Intersection => "\\bigcap",
        LargeOperator::BigUnion => "\\bigcup",
        LargeOperator::BigIntersection => "\\bigcap",
        LargeOperator::Limit => "\\lim",
        LargeOperator::Max => "\\max",
        LargeOperator::Min => "\\min",
        LargeOperator::Supremum => "\\sup",
        LargeOperator::Infimum => "\\inf",
        LargeOperator::ArgMax => "\\arg\\max",
        LargeOperator::ArgMin => "\\arg\\min",
    }
}

/// Convert accent to LaTeX string
pub fn accent_to_latex(accent: AccentType) -> &'static str {
    match accent {
        AccentType::Hat => "\\hat",
        AccentType::Check => "\\check",
        AccentType::Tilde => "\\tilde",
        AccentType::Acute => "\\acute",
        AccentType::Grave => "\\grave",
        AccentType::Dot => "\\dot",
        AccentType::DoubleDot => "\\ddot",
        AccentType::TripleDot => "\\dddot",
        AccentType::Bar => "\\bar",
        AccentType::Breve => "\\breve",
        AccentType::Vec => "\\vec",
    }
}

/// Convert space to LaTeX string
pub fn space_to_latex(space: SpaceType) -> &'static str {
    match space {
        SpaceType::Thin => "\\,",
        SpaceType::Medium => "\\:",
        SpaceType::Thick => "\\;",
        SpaceType::Quad => "\\quad",
        SpaceType::QQuad => "\\qquad",
        SpaceType::Negative => "\\!",
    }
}

/// Convert style to LaTeX string
pub fn style_to_latex(style: StyleType) -> &'static str {
    match style {
        StyleType::Normal => "\\mathrm",
        StyleType::Bold => "\\mathbf",
        StyleType::Italic => "\\mathit",
        StyleType::BoldItalic => "\\bm",
        StyleType::SansSerif => "\\mathsf",
        StyleType::SansSerifBold => "\\mathsf",
        StyleType::SansSerifItalic => "\\mathsf",
        StyleType::SansSerifBoldItalic => "\\mathsf",
        StyleType::Monospace => "\\mathtt",
        StyleType::Script => "\\mathcal",
        StyleType::BoldScript => "\\mathcal",
        StyleType::Fraktur => "\\mathfrak",
        StyleType::BoldFraktur => "\\mathfrak",
        StyleType::DoubleStruck => "\\mathbb",
    }
}

/// Check if a function name is a standard LaTeX function
pub fn is_standard_function(name: &str) -> bool {
    matches!(
        name,
        "sin"
            | "cos"
            | "tan"
            | "cot"
            | "sec"
            | "csc"
            | "sinh"
            | "cosh"
            | "tanh"
            | "coth"
            | "arcsin"
            | "arccos"
            | "arctan"
            | "log"
            | "ln"
            | "lg"
            | "exp"
            | "max"
            | "min"
            | "sup"
            | "inf"
            | "lim"
            | "limsup"
            | "liminf"
            | "det"
            | "dim"
            | "deg"
            | "gcd"
            | "hom"
            | "ker"
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_operator_to_latex() {
        assert_eq!(operator_to_latex(Operator::Plus), "+");
        assert_eq!(operator_to_latex(Operator::Minus), "-");
        assert_eq!(operator_to_latex(Operator::Multiply), "\\cdot");
        assert_eq!(operator_to_latex(Operator::Divide), "\\div");
        assert_eq!(operator_to_latex(Operator::Equals), "=");
        assert_eq!(operator_to_latex(Operator::NotEquals), "\\neq");
        assert_eq!(operator_to_latex(Operator::LessThan), "<");
        assert_eq!(operator_to_latex(Operator::GreaterThan), ">");
        assert_eq!(operator_to_latex(Operator::LessThanOrEqual), "\\leq");
        assert_eq!(operator_to_latex(Operator::GreaterThanOrEqual), "\\geq");
        assert_eq!(operator_to_latex(Operator::PlusMinus), "\\pm");
        assert_eq!(operator_to_latex(Operator::Times), "\\times");
        assert_eq!(operator_to_latex(Operator::Infinity), "\\infty");
        assert_eq!(operator_to_latex(Operator::Partial), "\\partial");
        assert_eq!(operator_to_latex(Operator::Nabla), "\\nabla");
    }

    #[test]
    fn test_fence_to_latex() {
        assert_eq!(fence_to_latex(Fence::Paren, true), "\\left(");
        assert_eq!(fence_to_latex(Fence::Paren, false), "\\right)");
        assert_eq!(fence_to_latex(Fence::Bracket, true), "\\left[");
        assert_eq!(fence_to_latex(Fence::Bracket, false), "\\right]");
        assert_eq!(fence_to_latex(Fence::Brace, true), "\\left\\{");
        assert_eq!(fence_to_latex(Fence::Brace, false), "\\right\\}");
        assert_eq!(fence_to_latex(Fence::Pipe, true), "\\left|");
        assert_eq!(fence_to_latex(Fence::Pipe, false), "\\right|");
    }

    #[test]
    fn test_large_operator_to_latex() {
        assert_eq!(large_operator_to_latex(LargeOperator::Sum), "\\sum");
        assert_eq!(large_operator_to_latex(LargeOperator::Product), "\\prod");
        assert_eq!(large_operator_to_latex(LargeOperator::Integral), "\\int");
        assert_eq!(
            large_operator_to_latex(LargeOperator::DoubleIntegral),
            "\\iint"
        );
        assert_eq!(
            large_operator_to_latex(LargeOperator::ContourIntegral),
            "\\oint"
        );
        assert_eq!(large_operator_to_latex(LargeOperator::Limit), "\\lim");
        assert_eq!(large_operator_to_latex(LargeOperator::Max), "\\max");
        assert_eq!(large_operator_to_latex(LargeOperator::Min), "\\min");
    }

    #[test]
    fn test_accent_to_latex() {
        assert_eq!(accent_to_latex(AccentType::Hat), "\\hat");
        assert_eq!(accent_to_latex(AccentType::Tilde), "\\tilde");
        assert_eq!(accent_to_latex(AccentType::Bar), "\\bar");
        assert_eq!(accent_to_latex(AccentType::Vec), "\\vec");
        assert_eq!(accent_to_latex(AccentType::Dot), "\\dot");
        assert_eq!(accent_to_latex(AccentType::DoubleDot), "\\ddot");
    }

    #[test]
    fn test_space_to_latex() {
        assert_eq!(space_to_latex(SpaceType::Thin), "\\,");
        assert_eq!(space_to_latex(SpaceType::Medium), "\\:");
        assert_eq!(space_to_latex(SpaceType::Thick), "\\;");
        assert_eq!(space_to_latex(SpaceType::Quad), "\\quad");
        assert_eq!(space_to_latex(SpaceType::Negative), "\\!");
    }

    #[test]
    fn test_style_to_latex() {
        assert_eq!(style_to_latex(StyleType::Normal), "\\mathrm");
        assert_eq!(style_to_latex(StyleType::Bold), "\\mathbf");
        assert_eq!(style_to_latex(StyleType::Italic), "\\mathit");
        assert_eq!(style_to_latex(StyleType::Monospace), "\\mathtt");
        assert_eq!(style_to_latex(StyleType::Script), "\\mathcal");
        assert_eq!(style_to_latex(StyleType::Fraktur), "\\mathfrak");
        assert_eq!(style_to_latex(StyleType::DoubleStruck), "\\mathbb");
    }

    #[test]
    fn test_is_standard_function() {
        assert!(is_standard_function("sin"));
        assert!(is_standard_function("cos"));
        assert!(is_standard_function("tan"));
        assert!(is_standard_function("log"));
        assert!(is_standard_function("ln"));
        assert!(is_standard_function("exp"));
        assert!(is_standard_function("max"));
        assert!(is_standard_function("min"));
        assert!(is_standard_function("lim"));
        assert!(!is_standard_function("foo"));
        assert!(!is_standard_function("bar"));
    }
}
